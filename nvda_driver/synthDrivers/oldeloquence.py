# oldeloquence.py — Old ETI-Eloquence 3.3 NVDA synth driver
#
# Uses eloquence_wrapper.dll via _oldeloq.py for audio capture.
# Wrapper serializes audio + indexes into a pull stream, eliminating
# the race condition that caused speech to freeze on rapid commands.

import re
import os
import queue
import threading
from collections import OrderedDict

import synthDriverHandler
from synthDriverHandler import SynthDriver, VoiceInfo
from autoSettingsUtils.driverSetting import NumericDriverSetting, BooleanDriverSetting, DriverSetting
from autoSettingsUtils.utils import StringParameterInfo
from speech.commands import IndexCommand, PitchCommand

from . import _oldeloq33 as _oldeloq

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

# ECI language IDs (param 9) — maps 3-letter code to (id, display_name)
langs = {
	'enu': (65536,  'American English'),
	'esp': (131072, 'Castilian Spanish'),
	'esm': (131073, 'Latin American Spanish'),
	'ptb': (458752, 'Brazilian Portuguese'),
	'fra': (196608, 'French'),
	'frc': (196609, 'French Canadian'),
	'eng': (65537,  'British English'),
	'ita': (327680, 'Italian'),
	'deu': (262144, 'German'),
}

# Locale mapping for NVDA
_LANG_TO_LOCALE = {
	'enu': 'en', 'eng': 'en', 'esp': 'es', 'esm': 'es',
	'fra': 'fr', 'frc': 'fr', 'ita': 'it', 'deu': 'de', 'ptb': 'pt',
}

variants = {
	1: "Wade", 2: "Flow", 3: "Bobby", 4: "Rocko",
	5: "Glen", 6: "Sandy", 7: "Grandma", 8: "Grandpa",
}

minRate = 0
maxRate = 100

# ---------------------------------------------------------------------------
# Text preprocessing
# ---------------------------------------------------------------------------

anticrash_res = {
	re.compile(r'\b(|\d+|\W+)(|un|anti|re)c(ae|\xe6)sur', re.I): r'\1\2seizur',
	re.compile(r"\b(|\d+|\W+)h'(r|v)[e]", re.I): r"\1h ' \2 e",
	re.compile(r'hesday'): ' hesday',
	re.compile(r"\b(|\d+|\W+)tz[s]che", re.I): r'\1tz sche',
}
pause_re = re.compile(r'([a-zA-Z])([.(),:;!?])( |$)')
time_re = re.compile(r"(\d):(\d+):(\d+)")


def _resub(dct, s):
	for r in dct:
		s = r.sub(dct[r], s)
	return s


# ---------------------------------------------------------------------------
# Background thread
# ---------------------------------------------------------------------------

class _BgThread(threading.Thread):
	def __init__(self, q, stop_event):
		super().__init__(daemon=True, name="OldEloquenceBgThread")
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
					"OldEloquence: bg thread error", exc_info=True)
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
		SynthDriver.VariantSetting(),
		SynthDriver.RateSetting(),
		SynthDriver.RateBoostSetting(),
		SynthDriver.PitchSetting(),
		SynthDriver.VolumeSetting(),
		SynthDriver.InflectionSetting(),
		NumericDriverSetting("hsz", "Head Size"),
		NumericDriverSetting("rgh", "Roughness"),
		NumericDriverSetting("bth", "Breathiness"),
		BooleanDriverSetting("ABRDICT", "Enable &abbreviation dictionary", False),
		DriverSetting("pauseMode", "Shorten &Pauses", defaultVal="2"),
	)

	description = 'Old ETI-Eloquence'
	name = 'oldeloquence'

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

		self.curvoice = "enu"
		self._variant = "1"
		self._ABRDICT = False
		self._pause_mode = 2
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
		# Index-only utterances: notify immediately
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
			safe = self._processText(raw)
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

		# Drop trailing empty blocks
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
			if self.speaking and indexesAfter:
				def cb(idxs=indexesAfter, g=gen):
					if self._speakGeneration == g:
						for i in idxs:
							synthDriverHandler.synthIndexReached.notify(
								synth=self, index=i)
				_oldeloq.feed_marker(on_done=cb)

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

	def _processText(self, text):
		"""Apply ECI text preprocessing: anticrash, volume, pauses, etc."""
		if not text:
			return ""
		text = text.encode('windows-1252', 'replace').decode('windows-1252', 'replace')
		text = _resub(anticrash_res, text)

		# Inline volume + sanitize backticks in user text
		vlm = self.getVParam(_oldeloq.VLM)
		text = " `vv%d %s" % (vlm, text.replace('`', ' '))

		# Fix times
		text = time_re.sub(r'\1:\2 \3', text)

		# Abbreviation dictionary
		if self._ABRDICT:
			text = "`da1 " + text
		else:
			text = "`da0 " + text

		# Shorten pauses
		if self._pause_mode == 2:
			text = pause_re.sub(r'\1 `p1\2\3', text)
		elif self._pause_mode == 1:
			if re.search(r'[.(),:;!?]$', text.strip()):
				text = re.sub(r'([.(),:;!?])$', r' `p1\1', text.strip())

		if not text.strip().endswith('`p2'):
			text += ' `p2'

		return text

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

	# Rate Boost — 2x speed via Sonic time-stretching (no pitch change)
	def _get_rateBoost(self):
		v = _oldeloq.dll_call("eloq_get_rate_boost")
		return v is not None and v > 100

	def _set_rateBoost(self, val):
		_oldeloq.dll_call("eloq_set_rate_boost", 200 if val else 100)

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

	# Inflection
	def _get_inflection(self):
		return self.getVParam(_oldeloq.FLUCTUATION)

	def _set_inflection(self, vl):
		self.setVParam(_oldeloq.FLUCTUATION, int(vl))

	# Head Size
	def _get_hsz(self):
		return self.getVParam(_oldeloq.HSZ)

	def _set_hsz(self, vl):
		self.setVParam(_oldeloq.HSZ, int(vl))

	# Roughness
	def _get_rgh(self):
		return self.getVParam(_oldeloq.RGH)

	def _set_rgh(self, vl):
		self.setVParam(_oldeloq.RGH, int(vl))

	# Breathiness
	def _get_bth(self):
		return self.getVParam(_oldeloq.BTH)

	def _set_bth(self, vl):
		self.setVParam(_oldeloq.BTH, int(vl))

	# Voice (language)
	def _getAvailableVoices(self):
		o = OrderedDict()
		engine_dir = os.path.join(os.path.dirname(__file__), "eloquence")
		if os.path.isdir(engine_dir):
			for name in os.listdir(engine_dir):
				if not name.lower().endswith('.syn'):
					continue
				lname = name.lower()[:-4]
				if lname in langs:
					info = langs[lname]
					locale = _LANG_TO_LOCALE.get(lname, 'en')
					o[str(info[0])] = VoiceInfo(str(info[0]), info[1], locale)
		return o

	def _get_voice(self):
		return str(self.curvoice)

	def _set_voice(self, vl):
		# Accept either language code ('enu') or numeric ID ('65536')
		if vl in langs:
			val = langs[vl][0]
		else:
			try:
				val = int(vl)
			except ValueError:
				val = 65536
		_oldeloq.dll_call("eloq_set_voice", val)
		self.curvoice = str(val)

	# Variant
	def _getAvailableVariants(self):
		return OrderedDict(
			(str(k), VoiceInfo(str(k), name))
			for k, name in variants.items()
		)

	def _get_variant(self):
		return self._variant

	def _set_variant(self, v):
		try:
			vv = int(v)
		except (ValueError, TypeError):
			vv = 1
		if vv not in variants:
			vv = 1
		self._variant = str(vv)
		_oldeloq.dll_call("eloq_set_variant", vv)
		# Re-apply rate after variant change
		if hasattr(self, '_rate'):
			self.setVParam(_oldeloq.RATE, self._rate)

	# Abbreviation dictionary
	def _get_ABRDICT(self):
		return self._ABRDICT

	def _set_ABRDICT(self, enable):
		self._ABRDICT = enable

	# Pause mode
	def _get_availablePausemodes(self):
		return {
			"0": StringParameterInfo("0", "Do not shorten"),
			"1": StringParameterInfo("1", "Shorten final punctuation"),
			"2": StringParameterInfo("2", "Shorten all punctuation"),
		}

	def _get_pauseMode(self):
		return str(self._pause_mode)

	def _set_pauseMode(self, val):
		try:
			pauseVal = int(val)
		except (ValueError, TypeError):
			pauseVal = 0
		if pauseVal not in (0, 1, 2):
			pauseVal = 0
		self._pause_mode = pauseVal


# ---------------------------------------------------------------------------
# 64-bit NVDA 2026.1+: bridge proxy
# ---------------------------------------------------------------------------
import ctypes as _ctypes
if _ctypes.sizeof(_ctypes.c_void_p) == 8:
	from _bridge.clients.synthDriverHost32.synthDriver import SynthDriverProxy32 as _Proxy32

	class SynthDriver(_Proxy32):
		name = "oldeloquence"
		description = "Old ETI-Eloquence"
		synthDriver32Path = os.path.dirname(__file__)
		synthDriver32Name = "oldeloquence"

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
