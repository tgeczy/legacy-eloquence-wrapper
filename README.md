# eloquence-wrapper

A C++ wrapper DLL for vintage ETI-Eloquence text-to-speech engines (versions 2.0 and 3.3), with NVDA synth driver included.

**Engine DLLs are NOT included.** You must supply your own Eloquence engine files.

## What this does

Eloquence 2.0 and 3.3 are old 32-bit TTS engines that either play audio directly to speakers (2.0) or deliver it through callbacks (3.3). This wrapper unifies both into a single pull-based API so modern screen readers can capture the audio and route it through their own audio pipeline.

- **Eloquence 3.3**: Audio via ECI callback. No hooks needed.
- **Eloquence 2.0**: Audio captured via MinHook waveOut interception (the engine resolves waveOut dynamically and plays directly to speakers).
- **Unified pull API**: `eloq_speak()` queues synthesis, `eloq_read()` returns audio/index/done items.
- **Silence trimming**: Caps consecutive silence at 60ms to reduce pauses between phrases.
- **Rate boost**: Sonic WSOLA time-stretching for speed beyond the engine's 100% ceiling without pitch change.

## API

```c
int  eloq_init(const wchar_t* engineDir);  // Load engine from directory
void eloq_free(void);                       // Release engine
int  eloq_version(void);                    // Returns 20 or 33
int  eloq_format(int* rate, int* bits, int* channels);

int  eloq_speak(const char* text);          // Queue text for synthesis
int  eloq_stop(void);                       // Cancel current speech
int  eloq_read(void* buf, int maxBytes, int* itemType, int* value);

int  eloq_set_variant(int variant);         // 1-8
int  eloq_set_voice(int voiceId);           // Language (3.3 only)
int  eloq_set_vparam(int param, int val);   // ECI voice params
int  eloq_get_vparam(int param);
int  eloq_set_rate_boost(int percent);      // 100=normal, 200=2x
int  eloq_get_rate_boost(void);
int  eloq_load_dict(const char* main, const char* root);
```

`eloq_read()` item types: `AUDIO` (buf filled), `INDEX` (value = index), `DONE`, `ERROR`, `NONE` (no data yet).

## Building

Requires MSVC with 32-bit target (the Eloquence engines are 32-bit).

```
cmake -B build -G Ninja -DCMAKE_BUILD_TYPE=Release
cmake --build build
```

**Important**: Always use `Release` build. `MinSizeRel` breaks MinHook's waveOut hooks.

## NVDA synth drivers

The `nvda_driver/` directory contains Python synth drivers for NVDA:

| File | Engine | Description |
|------|--------|-------------|
| `oldeloquence.py` | 3.3 | Multi-language, 8 voices, full ECI settings |
| `eloq20.py` | 2.0 | English-only, 8 variants |
| `_oldeloq.py` | Shared | Wrapper client, audio pipeline, ctypes bindings |

### Installation

1. Build `eloquence_wrapper.dll` (see above)
2. Place the DLL and driver `.py` files in your NVDA addon's `synthDrivers/` directory
3. Place your Eloquence engine files in a subdirectory:
   - 3.3: `synthDrivers/eloquence/` (needs `eci32d.dll` + language `.SYN` files)
   - 2.0: `synthDrivers/eloquence/` (needs `eci32d.dll` + `ENGSYN32.DLL`)
4. Restart NVDA and select the synth

### Features

- Capital pitch distinction (pitch rises on uppercase)
- Rate boost via Sonic time-stretching
- NVDA 2026.1+ 64-bit bridge compatible
- Say-all with proper index tracking
- Pause/resume support

## Third-party licenses

This project includes:

- **[MinHook](https://github.com/TsudaKageworthy/minhook)** - Minimalistic API hooking library (BSD 2-Clause). See [licenses/LICENSE-MinHook.txt](licenses/LICENSE-MinHook.txt).
- **[Sonic](https://github.com/nicknash/sonic)** - WSOLA time-stretching library (Apache 2.0). See [licenses/LICENSE-Sonic.txt](licenses/LICENSE-Sonic.txt).

## License

MIT. See [LICENSE](LICENSE).
