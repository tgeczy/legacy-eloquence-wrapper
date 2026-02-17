"""Shared wrapper client for vintage ETI-Eloquence (2.0 and 3.3).

Loads eloquence_wrapper.dll directly via ctypes.  On 64-bit NVDA 2026.1+
the entire synth driver (including this module) runs inside NVDA's built-in
32-bit bridge host, so this code always operates in a 32-bit process.
"""
from __future__ import annotations

import ctypes
import logging
import os
import queue
import threading
import time
from typing import Any, Callable, Dict, Optional, Tuple

LOGGER = logging.getLogger(__name__)

# Stream item types from eloquence_wrapper.cpp
ELOQ_ITEM_NONE = 0
ELOQ_ITEM_AUDIO = 1
ELOQ_ITEM_INDEX = 2
ELOQ_ITEM_DONE = 3
ELOQ_ITEM_ERROR = 4

# ECI voice param IDs
HSZ = 1
PITCH = 2
FLUCTUATION = 3
RGH = 4
BTH = 5
RATE = 6
VLM = 7

# DLL functions safe to call from Python
_ALLOWED_DLL_CALLS = frozenset({
	"eloq_set_vparam", "eloq_set_variant", "eloq_set_voice",
	"eloq_load_dict", "eloq_get_vparam",
	"eloq_set_rate_boost", "eloq_get_rate_boost",
})

AudioChunk = Tuple[bytes, Optional[Any], bool, int]


# ---------------------------------------------------------------------------
# AudioWorker (pulls audio from queue and feeds to nvwave.WavePlayer)
# ---------------------------------------------------------------------------

class AudioWorker(threading.Thread):
	"""Pulls audio events from the queue and feeds them to nvwave.WavePlayer."""

	def __init__(self, player, audio_queue: "queue.Queue[Optional[AudioChunk]]",
				 get_sequence: Callable[[], int],
				 player_lock: Optional[threading.RLock] = None,
				 auto_idle: bool = True):
		super().__init__(daemon=True, name="EloquenceAudioWorker")
		self._player = player
		self._queue = audio_queue
		self._get_sequence = get_sequence
		self._running = True
		self._stopping = False
		self._player_lock = player_lock or threading.RLock()
		self._auto_idle = auto_idle

	def run(self) -> None:
		while self._running:
			try:
				chunk = self._queue.get(timeout=0.1)
			except queue.Empty:
				continue
			if chunk is None:
				break

			data, second, is_final, seq = chunk

			if seq < self._get_sequence():
				self._queue.task_done()
				continue

			# Marker callback: feed empty buffer with onDone to WavePlayer
			if callable(second):
				if not self._stopping:
					try:
						with self._player_lock:
							if not self._stopping and self._player:
								self._player.feed(b"", onDone=second)
					except Exception:
						LOGGER.exception("Marker feed failed")
				self._queue.task_done()
				continue

			# Idle signal
			if second == "IDLE":
				if not self._stopping:
					try:
						with self._player_lock:
							if not self._stopping and self._player:
								self._player.idle()
					except Exception:
						LOGGER.exception("Player idle failed")
				self._queue.task_done()
				continue

			# Done marker (from _read_loop)
			if not data and second is None:
				if is_final and self._auto_idle:
					with self._player_lock:
						if not self._stopping:
							self._player.idle()
					if not self._stopping:
						self._invoke_done_callback()
				self._queue.task_done()
				continue

			if self._stopping:
				self._queue.task_done()
				continue

			try:
				with self._player_lock:
					if not self._stopping and self._player:
						self._player.feed(data)
			except FileNotFoundError:
				LOGGER.warning("Sound device not found during feed")
			except Exception:
				LOGGER.exception("WavePlayer feed failed")
			self._queue.task_done()

	def stop(self) -> None:
		self._stopping = True
		self._running = False
		self._queue.put(None)

	def _invoke_done_callback(self) -> None:
		if _on_done:
			try:
				_on_done()
			except Exception:
				LOGGER.exception("Done callback failed")


# ---------------------------------------------------------------------------
# EloquenceClient (direct ctypes access to eloquence_wrapper.dll)
# ---------------------------------------------------------------------------

class EloquenceClient:
	"""Direct ctypes access to eloquence_wrapper.dll."""

	def __init__(self) -> None:
		self._dll = None
		self._audio_queue: "queue.Queue[Optional[AudioChunk]]" = queue.Queue()
		self._player = None
		self._player_lock = threading.RLock()
		self._audio_worker: Optional[AudioWorker] = None
		self._should_stop = False
		self._sequence = 0
		self._current_seq = 0
		# Audio format
		self._sample_rate = 0
		self._channels = 0
		self._bits_per_sample = 0
		# Engine tracking
		self._engine_dir: Optional[str] = None
		# Read buffer
		self._buf_size = 65536
		self._audio_buf = None
		self._out_type = None
		self._out_value = None

	def do_initialize(self, dll_path: str, engine_dir: str) -> Dict[str, Any]:
		"""Load the wrapper DLL and initialize the ECI engine."""
		if self._engine_dir == engine_dir and self._dll is not None:
			# Same engine already initialized — reuse.
			LOGGER.info("_oldeloq: reusing existing engine (dir=%s)", engine_dir)
			return {
				"format": {
					"sampleRate": self._sample_rate,
					"channels": self._channels,
					"bitsPerSample": self._bits_per_sample,
				},
			}

		if self._dll is None:
			LOGGER.info("_oldeloq: loading wrapper DLL from %s", dll_path)
			self._dll = ctypes.cdll.LoadLibrary(dll_path)
			self._setup_ctypes()

		# If switching engines, free the old one first.
		if self._engine_dir is not None and self._engine_dir != engine_dir:
			self._dll.eloq_free()
			self._engine_dir = None

		self._audio_buf = ctypes.create_string_buffer(self._buf_size)
		self._out_type = ctypes.c_int(0)
		self._out_value = ctypes.c_int(0)

		LOGGER.info("_oldeloq: calling eloq_init(%s)", engine_dir)
		rc = self._dll.eloq_init(engine_dir)
		if rc != 0:
			raise RuntimeError(f"eloq_init returned {rc}")
		self._engine_dir = engine_dir

		# Query audio format
		sr = ctypes.c_int(0)
		bps = ctypes.c_int(0)
		ch = ctypes.c_int(0)
		if self._dll.eloq_format(ctypes.byref(sr), ctypes.byref(bps), ctypes.byref(ch)) == 0:
			self._sample_rate = sr.value
			self._bits_per_sample = bps.value
			self._channels = ch.value
		else:
			self._sample_rate = 11025
			self._channels = 1
			self._bits_per_sample = 16

		ver = self._dll.eloq_version()
		LOGGER.info("_oldeloq: init OK — version=%d, format=%dHz/%dbit/%dch",
					ver, self._sample_rate, self._bits_per_sample, self._channels)

		return {
			"format": {
				"sampleRate": self._sample_rate,
				"channels": self._channels,
				"bitsPerSample": self._bits_per_sample,
			},
		}

	def _setup_ctypes(self):
		dll = self._dll
		dll.eloq_init.argtypes = (ctypes.c_wchar_p,)
		dll.eloq_init.restype = ctypes.c_int
		dll.eloq_free.argtypes = ()
		dll.eloq_free.restype = None
		dll.eloq_version.argtypes = ()
		dll.eloq_version.restype = ctypes.c_int
		dll.eloq_format.argtypes = (
			ctypes.POINTER(ctypes.c_int),
			ctypes.POINTER(ctypes.c_int),
			ctypes.POINTER(ctypes.c_int),
		)
		dll.eloq_format.restype = ctypes.c_int
		dll.eloq_speak.argtypes = (ctypes.c_char_p,)
		dll.eloq_speak.restype = ctypes.c_int
		dll.eloq_stop.argtypes = ()
		dll.eloq_stop.restype = ctypes.c_int
		dll.eloq_read.argtypes = (
			ctypes.c_void_p,
			ctypes.c_int,
			ctypes.POINTER(ctypes.c_int),
			ctypes.POINTER(ctypes.c_int),
		)
		dll.eloq_read.restype = ctypes.c_int
		dll.eloq_set_variant.argtypes = (ctypes.c_int,)
		dll.eloq_set_variant.restype = ctypes.c_int
		dll.eloq_set_vparam.argtypes = (ctypes.c_int, ctypes.c_int)
		dll.eloq_set_vparam.restype = ctypes.c_int
		dll.eloq_get_vparam.argtypes = (ctypes.c_int,)
		dll.eloq_get_vparam.restype = ctypes.c_int
		dll.eloq_set_voice.argtypes = (ctypes.c_int,)
		dll.eloq_set_voice.restype = ctypes.c_int
		dll.eloq_load_dict.argtypes = (ctypes.c_char_p, ctypes.c_char_p)
		dll.eloq_load_dict.restype = ctypes.c_int
		dll.eloq_set_rate_boost.argtypes = (ctypes.c_int,)
		dll.eloq_set_rate_boost.restype = ctypes.c_int
		dll.eloq_get_rate_boost.argtypes = ()
		dll.eloq_get_rate_boost.restype = ctypes.c_int

	# ------------------------------------------------------------------
	# Audio
	def initialize_audio(self, channels: int, sample_rate: int, bits_per_sample: int) -> None:
		if self._player:
			return
		import nvwave
		import config
		try:
			from buildVersion import version_year
		except ImportError:
			version_year = 2025

		if version_year >= 2025:
			device = config.conf["audio"]["outputDevice"]
			player = nvwave.WavePlayer(channels, sample_rate, bits_per_sample,
									   outputDevice=device)
		else:
			device = config.conf["speech"]["outputDevice"]
			player = nvwave.WavePlayer(channels, sample_rate, bits_per_sample,
									   outputDevice=device, buffered=True)
		self._player = player
		self._audio_worker = AudioWorker(player, self._audio_queue,
										 lambda: self._sequence,
										 player_lock=self._player_lock,
										 auto_idle=False)
		self._audio_worker.start()

	# ------------------------------------------------------------------
	# Speech (blocking)
	def do_speak(self, text_bytes: bytes) -> bool:
		"""Start speech and pump read loop. Returns True on success."""
		self._should_stop = False
		self._current_seq = self._sequence
		rc = self._dll.eloq_speak(text_bytes)
		if rc != 0:
			LOGGER.error("eloq_speak returned %d for %r", rc, text_bytes[:80])
			self._audio_queue.put((b"", None, True, self._current_seq))
			return False
		LOGGER.debug("eloq_speak OK, seq=%d, entering read loop", self._current_seq)
		return self._read_loop()

	def _read_loop(self) -> bool:
		"""Poll eloq_read() and push audio to queue. Returns True if completed normally."""
		audio_chunks = 0
		audio_bytes = 0
		while not self._should_stop:
			try:
				n = self._dll.eloq_read(
					self._audio_buf,
					self._buf_size,
					ctypes.byref(self._out_type),
					ctypes.byref(self._out_value),
				)
			except Exception:
				LOGGER.exception("eloq_read crashed")
				self._audio_queue.put((b"", None, True, self._current_seq))
				return False

			t = self._out_type.value

			if t == ELOQ_ITEM_AUDIO and n > 0:
				audio_chunks += 1
				audio_bytes += n
				self._audio_queue.put((bytes(self._audio_buf.raw[:n]), None, False, self._current_seq))
			elif t == ELOQ_ITEM_INDEX:
				# Index from engine — ignored (we use feed_marker instead)
				pass
			elif t == ELOQ_ITEM_DONE:
				LOGGER.debug("read_loop DONE: %d chunks, %d bytes", audio_chunks, audio_bytes)
				self._audio_queue.put((b"", None, True, self._current_seq))
				return True
			elif t == ELOQ_ITEM_ERROR:
				LOGGER.error("Wrapper error %d", self._out_value.value)
				self._audio_queue.put((b"", None, True, self._current_seq))
				return False
			elif t == ELOQ_ITEM_NONE:
				time.sleep(0.001)
		LOGGER.debug("read_loop stopped: %d chunks, %d bytes", audio_chunks, audio_bytes)
		return False

	# ------------------------------------------------------------------
	# Control
	def dll_call(self, func_name: str, *args):
		if func_name not in _ALLOWED_DLL_CALLS:
			raise ValueError(f"Disallowed: {func_name}")
		fn = getattr(self._dll, func_name, None)
		if fn:
			return fn(*args)

	def stop(self) -> None:
		self._sequence += 1
		self._should_stop = True
		if self._audio_worker:
			self._audio_worker._stopping = True
		if self._dll:
			self._dll.eloq_stop()
		# Call player.stop() WITHOUT the lock — WavePlayer.stop() is
		# thread-safe.  Using the lock here would deadlock if AudioWorker
		# is blocked inside feed() on WASAPI buffer space.
		if self._player:
			try:
				self._player.stop()
			except Exception:
				LOGGER.exception("WavePlayer stop failed")

	def pause(self, switch: bool) -> None:
		with self._player_lock:
			if self._player:
				self._player.pause(switch)

	def feed_marker(self, on_done=None) -> None:
		self._audio_queue.put((b"", on_done, False, self._current_seq))

	def player_idle(self) -> None:
		self._audio_queue.put((b"", "IDLE", False, self._current_seq))

	def get_format(self) -> Dict[str, int]:
		return {
			"sampleRate": self._sample_rate,
			"channels": self._channels,
			"bitsPerSample": self._bits_per_sample,
		}

	# ------------------------------------------------------------------
	# Shutdown
	def shutdown(self) -> None:
		if self._audio_worker:
			self._audio_worker.stop()
			self._audio_worker.join(timeout=1)
			self._audio_worker = None
		with self._player_lock:
			if self._player:
				self._player.close()
				self._player = None
		if self._dll:
			self._dll.eloq_free()
			import ctypes
			ctypes.windll.kernel32.FreeLibrary(self._dll._handle)
			self._dll = None
		self._engine_dir = None
		self._audio_queue = queue.Queue()
		self._sequence = 0
		self._current_seq = 0


# ---------------------------------------------------------------------------
# Module-level singleton and public API
# ---------------------------------------------------------------------------

_client: EloquenceClient = EloquenceClient()
_on_done: Optional[Callable] = None
_format: Dict[str, int] = {}


def initialize(engine_dir: str, done_callback=None) -> Dict[str, Any]:
	"""Load the wrapper DLL and initialize the engine."""
	global _on_done, _format
	_on_done = done_callback

	addon_dir = os.path.abspath(os.path.dirname(__file__))
	dll_path = os.path.join(addon_dir, "eloquence_wrapper.dll")

	result = _client.do_initialize(dll_path, engine_dir)

	_format = result.get("format", {})

	_client.initialize_audio(
		channels=_format.get("channels", 1),
		sample_rate=_format.get("sampleRate", 11025),
		bits_per_sample=_format.get("bitsPerSample", 16),
	)

	# Load dictionaries if available (3.3 only, no-op on 2.0)
	main_dic = os.path.join(engine_dir, "main.dic")
	root_dic = os.path.join(engine_dir, "root.dic")
	main_path = main_dic.encode("mbcs") if os.path.exists(main_dic) else None
	root_path = root_dic.encode("mbcs") if os.path.exists(root_dic) else None
	if main_path or root_path:
		try:
			_client.dll_call("eloq_load_dict", main_path, root_path)
		except Exception:
			LOGGER.exception("Failed to load dictionaries")

	return result


def speak(text: str) -> bool:
	"""Speak text (blocks until audio is queued). Returns True on success."""
	# Reset AudioWorker's stopping flag so it will feed new audio
	if _client._audio_worker:
		_client._audio_worker._stopping = False
	text_bytes = text.encode("windows-1252", errors="replace")
	LOGGER.debug("speak: %r", text_bytes[:120])
	return _client.do_speak(text_bytes)


def dll_call(func_name: str, *args):
	"""Call an eloq_* function by name."""
	return _client.dll_call(func_name, *args)


def stop() -> None:
	_client.stop()


def pause(switch: bool) -> None:
	_client.pause(switch)


def feed_marker(on_done=None) -> None:
	"""Feed an empty marker to the player. on_done fires when it reaches playback."""
	_client.feed_marker(on_done)


def player_idle() -> None:
	"""Signal the player that no more audio is coming."""
	_client.player_idle()


def terminate() -> None:
	_client.shutdown()


def get_format() -> Dict[str, int]:
	return dict(_format)


def check(engine_dir: str) -> bool:
	"""Check if wrapper DLL and engine DLL are present."""
	addon_dir = os.path.abspath(os.path.dirname(__file__))
	wrapper_dll = os.path.join(addon_dir, "eloquence_wrapper.dll")
	eci_dll = os.path.join(engine_dir, "eci32d.dll")
	return os.path.isfile(wrapper_dll) and os.path.isfile(eci_dll)
