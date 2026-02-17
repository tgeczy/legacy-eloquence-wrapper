# eloq20.py — ETI-Eloquence 2.0 NVDA synth driver
#
# Uses eloquence_wrapper.dll via _oldeloq.py for audio capture.
# The wrapper hooks waveOut via MinHook to intercept audio that
# Eloquence 2.0's ENGSYN32.DLL plays directly to speakers.

import os
import queue
import threading
from collections import OrderedDict

import synthDriverHandler
from synthDriverHandler import SynthDriver, VoiceInfo
from speech.commands import IndexCommand, PitchCommand

from . import _oldeloq20 as _oldeloq

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

variants = {
	1: "Eddy", 2: "Flo", 3: "Bobbie", 4: "Rocko",
	5: "Glen", 6: "Sandy", 7: "Grandma", 8: "Grandpa",
}

minRate = 0
maxRate = 100


# ---------------------------------------------------------------------------
# Background thread
# ---------------------------------------------------------------------------

class _BgThread(threading.Thread):
	def __init__(self, q, stop_event):
		super().__init__(daemon=True, name="Eloq20BgThread")
		self._q = q
		self._stop = stop_event

	def run(self):
		while not self._stop.is_set():
			try:
				item = self._q.get(timeout=0.2)
			except queue.Empty:
				continue
			try:
				if item is None:
					return
				func, args, kwargs = item
				func(*args, **kwargs)
			except Exception:
				import logging
				logging.getLogger(__name__).error(
					"Eloq20: bg thread error", exc_info=True)
			finally:
				try:
					self._q.task_done()
				except Exception:
					pass


# ---------------------------------------------------------------------------
# SynthDriver
# ---------------------------------------------------------------------------

class SynthDriver(synthDriverHandler.SynthDriver):
	supportedCommands = {IndexCommand, PitchCommand}
	supportedNotifications = {
		synthDriverHandler.synthIndexReached,
		synthDriverHandler.synthDoneSpeaking,
	}
	supportedSettings = (
		SynthDriver.VoiceSetting(),
		SynthDriver.RateSetting(),
		SynthDriver.RateBoostSetting(),
		SynthDriver.PitchSetting(),
		SynthDriver.VolumeSetting(),
	)

	description = 'Old ETI-Eloquence 2.0'
	name = 'eloq20'

	@classmethod
	def check(cls):
		engine_dir = os.path.join(os.path.dirname(__file__), "eloquence")
		return _oldeloq.check(engine_dir)

	def __init__(self):
		# Ensure config.pre_configSave exists (bridge host compat)
		import config
		if not hasattr(config, 'pre_configSave'):
			import extensionPoints
			config.pre_configSave = extensionPoints.Action()
		super().__init__()

		engine_dir = os.path.join(os.path.dirname(__file__), "eloquence")
		_oldeloq.initialize(engine_dir)

		self._voice = "1"
		self._basePitch = 50

		self.speaking = False
		self._speakGeneration = 0
		self._terminating = False

		self._bgQueue = queue.Queue()
		self._bgStop = threading.Event()
		self._bgThread = _BgThread(self._bgQueue, self._bgStop)
		self._bgThread.start()

		# Apply initial rate
		self.rate = 50

	def _enqueue(self, func, *args, **kwargs):
		if not self._terminating:
			self._bgQueue.put((func, args, kwargs))

	def terminate(self):
		self._terminating = True
		self.cancel()
		try:
			self._bgStop.set()
			self._bgQueue.put(None)
			self._bgThread.join(timeout=2.0)
		except Exception:
			pass
		_oldeloq.terminate()

	def cancel(self):
		self._speakGeneration += 1
		self.speaking = False
		_oldeloq.stop()
		# Restore base pitch after capital pitch changes
		self.setVParam(_oldeloq.PITCH, self._basePitch)
		try:
			while True:
				self._bgQueue.get_nowait()
				self._bgQueue.task_done()
		except queue.Empty:
			pass

	def pause(self, switch):
		_oldeloq.pause(switch)

	# --- Speaking ---

	def speak(self, speechSequence):
		if len(speechSequence) == 1 and isinstance(speechSequence[0], IndexCommand):
			self._enqueue(self._notifyIndexesAndDone, [speechSequence[0].index])
			return

		blocks, anyText, allIndexes = self._buildBlocks(speechSequence)
		if not anyText:
			self._enqueue(self._notifyIndexesAndDone, allIndexes)
			return
		self._enqueue(self._speakBg, blocks)

	def _buildBlocks(self, speechSequence):
		blocks = []
		textBuf = []
		pendingIndexes = []
		curPitchOffset = 0

		def flush():
			raw = " ".join(textBuf)
			textBuf.clear()
			# Eloquence 2.0 is English-only — strip non-ASCII to avoid
			# engine crashes on accented/Unicode characters.
			safe = ''.join(ch if ord(ch) < 128 else ' ' for ch in raw)
			safe = safe.replace('`', ' ').strip()
			blocks.append((safe, pendingIndexes.copy(), curPitchOffset))
			pendingIndexes.clear()

		for item in speechSequence:
			if isinstance(item, str):
				textBuf.append(item)
			elif isinstance(item, IndexCommand):
				pendingIndexes.append(item.index)
			elif isinstance(item, PitchCommand):
				if textBuf:
					flush()
				curPitchOffset = item.offset

		if textBuf or pendingIndexes:
			flush()

		while blocks and not blocks[-1][0] and not blocks[-1][1]:
			blocks.pop()

		anyText = any(bool(t) for (t, _, __) in blocks)
		allIndexes = []
		for (_, idxs, __) in blocks:
			allIndexes.extend(idxs)
		return blocks, anyText, allIndexes

	def _notifyIndexesAndDone(self, indexes):
		for i in indexes:
			synthDriverHandler.synthIndexReached.notify(synth=self, index=i)
		synthDriverHandler.synthDoneSpeaking.notify(synth=self)
		self.speaking = False

	def _speakBg(self, blocks):
		self._speakGeneration += 1
		gen = self._speakGeneration
		self.speaking = True
		basePitch = self._basePitch
		lastPitch = None

		for (text, indexesAfter, pitchOffset) in blocks:
			if not self.speaking:
				break
			# Apply pitch offset for capital letter distinction
			pitch = max(0, min(100, int(basePitch + basePitch * pitchOffset / 100))) if pitchOffset else basePitch
			if pitch != lastPitch:
				self.setVParam(_oldeloq.PITCH, pitch)
				lastPitch = pitch
			if text:
				if not _oldeloq.speak(text):
					self.speaking = False
					break
			# Fire indexes right after synthesis so NVDA can prefetch
			# the next chunk while this one is still playing.
			if self.speaking and indexesAfter:
				if self._speakGeneration == gen:
					for i in indexesAfter:
						synthDriverHandler.synthIndexReached.notify(
							synth=self, index=i)

		# Restore base pitch
		if lastPitch != basePitch:
			self.setVParam(_oldeloq.PITCH, basePitch)

		if not self.speaking:
			synthDriverHandler.synthDoneSpeaking.notify(synth=self)
			return

		def doneCb(g=gen):
			if self._speakGeneration == g:
				self.speaking = False
				synthDriverHandler.synthDoneSpeaking.notify(synth=self)

		_oldeloq.feed_marker(on_done=doneCb)
		_oldeloq.player_idle()

	# --- Settings ---

	def _percentToParam(self, val, minVal, maxVal):
		return int(round(minVal + (maxVal - minVal) * float(val) / 100.0))

	def _paramToPercent(self, val, minVal, maxVal):
		if maxVal == minVal:
			return 0
		return int(round(100.0 * float(val - minVal) / float(maxVal - minVal)))

	def getVParam(self, pr):
		v = _oldeloq.dll_call("eloq_get_vparam", pr)
		return v if v is not None and v >= 0 else 0

	def setVParam(self, pr, vl):
		_oldeloq.dll_call("eloq_set_vparam", int(pr), int(vl))

	# Rate
	def _get_rate(self):
		return self._paramToPercent(self.getVParam(_oldeloq.RATE), minRate, maxRate)

	def _set_rate(self, vl):
		self._rate = self._percentToParam(vl, minRate, maxRate)
		self.setVParam(_oldeloq.RATE, self._rate)

	# Pitch
	def _get_pitch(self):
		return self._basePitch

	def _set_pitch(self, vl):
		self._basePitch = int(vl)
		self.setVParam(_oldeloq.PITCH, int(vl))

	# Volume
	def _get_volume(self):
		return self.getVParam(_oldeloq.VLM)

	def _set_volume(self, vl):
		self.setVParam(_oldeloq.VLM, int(vl))

	# Rate Boost — 2x speed via Sonic time-stretching (no pitch change)
	def _get_rateBoost(self):
		v = _oldeloq.dll_call("eloq_get_rate_boost")
		return v is not None and v > 100

	def _set_rateBoost(self, val):
		_oldeloq.dll_call("eloq_set_rate_boost", 200 if val else 100)

	# Voice — 8 ECI variants exposed as NVDA voices
	def _getAvailableVoices(self):
		return OrderedDict(
			(str(k), VoiceInfo(str(k), name, "en"))
			for k, name in variants.items()
		)

	def _get_voice(self):
		return self._voice

	def _set_voice(self, v):
		try:
			vv = int(v)
		except (ValueError, TypeError):
			vv = 1
		if vv not in variants:
			vv = 1
		self._voice = str(vv)
		_oldeloq.dll_call("eloq_set_variant", vv)
		# Re-apply rate after variant change
		if hasattr(self, '_rate'):
			self.setVParam(_oldeloq.RATE, self._rate)


# ---------------------------------------------------------------------------
# 64-bit NVDA 2026.1+: bridge proxy
# ---------------------------------------------------------------------------
import ctypes as _ctypes
if _ctypes.sizeof(_ctypes.c_void_p) == 8:
	from _bridge.clients.synthDriverHost32.synthDriver import SynthDriverProxy32 as _Proxy32

	class SynthDriver(_Proxy32):
		name = "eloq20"
		description = "Old ETI-Eloquence 2.0"
		synthDriver32Path = os.path.dirname(__file__)
		synthDriver32Name = "eloq20"

		_BRIDGE_SAFE = frozenset({"voice", "variant", "rate", "pitch", "volume", "rateBoost"})

		def _get_supportedSettings(self):
			return [s for s in super()._get_supportedSettings() if s.id in self._BRIDGE_SAFE]

		@classmethod
		def check(cls):
			if not super().check():
				return False
			base = os.path.dirname(__file__)
			return (
				os.path.isfile(os.path.join(base, "eloquence_wrapper.dll"))
				and os.path.isfile(os.path.join(base, "eloquence", "eci32d.dll"))
			)
