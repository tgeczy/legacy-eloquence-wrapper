// eloquence_wrapper.cpp
//
// C++ wrapper DLL for vintage ETI-Eloquence 2.0 and 3.3.
// - 3.3: Audio via ECI callback (msg==0) + output buffer. No hooks needed.
// - 2.0: Audio via MinHook waveOut interception (ENGSYN32.DLL resolves waveOut
//         dynamically via GetProcAddress and plays directly).
// - Exposes a unified pull API (eloq_read) so NVDA Python driver can feed NVWave.
//
// NOTE: Both Eloquence versions are 32-bit. Build this wrapper as 32-bit.

#define WIN32_LEAN_AND_MEAN
#define NOMINMAX

#ifndef ELOQ_API
#define ELOQ_API __declspec(dllexport)
#endif

// Stream item types.
#define ELOQ_ITEM_NONE  0
#define ELOQ_ITEM_AUDIO 1
#define ELOQ_ITEM_INDEX 2
#define ELOQ_ITEM_DONE  3
#define ELOQ_ITEM_ERROR 4

// Modes.
#define ELOQ_MODE_NONE 0
#define ELOQ_MODE_33   33
#define ELOQ_MODE_20   20

#include <windows.h>
#include <mmsystem.h>
#include <intrin.h>

#include <atomic>
#include <cstddef>
#include <cstdint>
#include <cstring>
#include <deque>
#include <mutex>
#include <string>
#include <thread>
#include <vector>

#include "MinHook.h"
#include "sonic.h"

#include <cstdio>

#pragma comment(lib, "user32.lib")

// ------------------------------------------------------------
// Debug tracing (writes to eloq_debug.log next to the DLL)
// ------------------------------------------------------------
static FILE* g_logFile = nullptr;
static std::mutex g_logMtx;

static void dbgOpen() {
    if (g_logFile) return;
    // Put log file next to the DLL itself.
    char dllPath[MAX_PATH] = {};
    HMODULE hSelf = nullptr;
    GetModuleHandleExA(
        GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
        (LPCSTR)&dbgOpen, &hSelf);
    if (hSelf) GetModuleFileNameA(hSelf, dllPath, MAX_PATH);
    // Replace DLL name with log name.
    char* slash = strrchr(dllPath, '\\');
    if (slash) strcpy(slash + 1, "eloq_debug.log");
    else strcpy(dllPath, "eloq_debug.log");
    g_logFile = fopen(dllPath, "w");
}

static void dbg(const char* fmt, ...) {
    std::lock_guard<std::mutex> lk(g_logMtx);
    if (!g_logFile) dbgOpen();
    if (!g_logFile) return;
    DWORD ms = GetTickCount();
    fprintf(g_logFile, "[%u.%03u] ", ms / 1000, ms % 1000);
    va_list ap;
    va_start(ap, fmt);
    vfprintf(g_logFile, fmt, ap);
    va_end(ap);
    fprintf(g_logFile, "\n");
    fflush(g_logFile);
}

// ------------------------------------------------------------
// ECI function pointer types (stdcall, 32-bit)
// ------------------------------------------------------------
typedef void* (__stdcall* eciNewFunc)(void);
typedef void  (__stdcall* eciDeleteFunc)(void* handle);
typedef void  (__stdcall* eciRequestLicenseFunc)(int key);
typedef int   (__stdcall* eciSetOutputBufferFunc)(void* handle, int samples, void* buffer);
typedef int   (__stdcall* eciSetOutputDeviceFunc)(void* handle, int device);
typedef int   (__stdcall* eciRegisterCallbackFunc)(void* handle, void* cb, void* data);
typedef int   (__stdcall* eciSetParamFunc)(void* handle, int param, int value);
typedef int   (__stdcall* eciGetParamFunc)(void* handle, int param);
typedef int   (__stdcall* eciSetVoiceParamFunc)(void* handle, int voice, int param, int value);
typedef int   (__stdcall* eciGetVoiceParamFunc)(void* handle, int voice, int param);
typedef int   (__stdcall* eciCopyVoiceFunc)(void* handle, int variant, int reserved);
typedef int   (__stdcall* eciAddTextFunc)(void* handle, const char* text);
typedef int   (__stdcall* eciInsertIndexFunc)(void* handle, int index);
typedef int   (__stdcall* eciSynthesizeFunc)(void* handle);
typedef int   (__stdcall* eciStopFunc)(void* handle);
typedef int   (__stdcall* eciSpeakingFunc)(void* handle);
typedef int   (__stdcall* eciSynchronizeFunc)(void* handle);
typedef int   (__stdcall* eciVersionFunc)(void* handle);
typedef int   (__stdcall* eciNewDictFunc)(void* handle);
typedef int   (__stdcall* eciSetDictFunc)(void* handle, int dict);
typedef int   (__stdcall* eciLoadDictFunc)(void* handle, int dict, int type, const char* path);

// ECI callback: int __cdecl callback(int handle, int msg, int length, void* data)
typedef int (__cdecl* EciCallbackFunc)(int, int, int, void*);

// ------------------------------------------------------------
// WinMM function pointer types + originals (for 2.0 hooks)
// ------------------------------------------------------------
typedef MMRESULT(WINAPI* waveOutOpenFunc)(LPHWAVEOUT, UINT, LPCWAVEFORMATEX, DWORD_PTR, DWORD_PTR, DWORD);
typedef MMRESULT(WINAPI* waveOutPrepareHeaderFunc)(HWAVEOUT, LPWAVEHDR, UINT);
typedef MMRESULT(WINAPI* waveOutWriteFunc)(HWAVEOUT, LPWAVEHDR, UINT);
typedef MMRESULT(WINAPI* waveOutUnprepareHeaderFunc)(HWAVEOUT, LPWAVEHDR, UINT);
typedef MMRESULT(WINAPI* waveOutResetFunc)(HWAVEOUT);
typedef MMRESULT(WINAPI* waveOutCloseFunc)(HWAVEOUT);

static waveOutOpenFunc            g_waveOutOpenOrig = nullptr;
static waveOutPrepareHeaderFunc   g_waveOutPrepareHeaderOrig = nullptr;
static waveOutWriteFunc           g_waveOutWriteOrig = nullptr;
static waveOutUnprepareHeaderFunc g_waveOutUnprepareHeaderOrig = nullptr;
static waveOutResetFunc           g_waveOutResetOrig = nullptr;
static waveOutCloseFunc           g_waveOutCloseOrig = nullptr;

// ------------------------------------------------------------
// Stream queue items
// ------------------------------------------------------------
struct StreamItem {
    int type = ELOQ_ITEM_NONE;
    int value = 0;
    uint32_t gen = 0;
    std::vector<uint8_t> data;
    size_t offset = 0;

    StreamItem() = default;
    StreamItem(int t, int v, uint32_t g) : type(t), value(v), gen(g), offset(0) {}
};

// ------------------------------------------------------------
// Command queue
// ------------------------------------------------------------
struct Cmd {
    enum Type { CMD_SPEAK, CMD_QUIT } type = CMD_SPEAK;
    uint32_t cancelSnapshot = 0;
    std::string text; // MBCS-encoded
};

// ------------------------------------------------------------
// Dirty-tracking settings
// ------------------------------------------------------------
struct SettingInt {
    std::atomic<int> value{ 0 };
    std::atomic<int> dirty{ 0 };
};

// ------------------------------------------------------------
// Global wrapper state
// ------------------------------------------------------------
struct ELOQ_STATE {
    // Mode
    int mode = ELOQ_MODE_NONE;

    // DLLs
    HMODULE cwlModule = nullptr;
    HMODULE eciModule = nullptr;
    HMODULE engsynModule = nullptr; // 2.0: auto-loaded via ECI32D's import

    // 2.0: Thread/handle-based hook filtering (replaces caller-module check).
    DWORD workerThreadId = 0;
    HWAVEOUT eloqWaveHandle = nullptr;

    // DLL directory (for CFG patching, dict loading)
    std::wstring dllDir;

    // ECI function pointers
    eciNewFunc              fnNew = nullptr;
    eciDeleteFunc           fnDelete = nullptr;
    eciRequestLicenseFunc   fnRequestLicense = nullptr;
    eciSetOutputBufferFunc  fnSetOutputBuffer = nullptr;
    eciSetOutputDeviceFunc  fnSetOutputDevice = nullptr;
    eciRegisterCallbackFunc fnRegisterCallback = nullptr;
    eciSetParamFunc         fnSetParam = nullptr;
    eciGetParamFunc         fnGetParam = nullptr;
    eciSetVoiceParamFunc    fnSetVoiceParam = nullptr;
    eciGetVoiceParamFunc    fnGetVoiceParam = nullptr;
    eciCopyVoiceFunc        fnCopyVoice = nullptr;
    eciAddTextFunc          fnAddText = nullptr;
    eciInsertIndexFunc      fnInsertIndex = nullptr;
    eciSynthesizeFunc       fnSynthesize = nullptr;
    eciStopFunc             fnStop = nullptr;
    eciSpeakingFunc         fnSpeaking = nullptr;
    eciSynchronizeFunc      fnSynchronize = nullptr;
    eciVersionFunc          fnVersion = nullptr;
    eciNewDictFunc          fnNewDict = nullptr;
    eciSetDictFunc          fnSetDict = nullptr;
    eciLoadDictFunc         fnLoadDict = nullptr;

    // ECI handle
    void* handle = nullptr;
    int dictHandle = -1;

    // 3.3: output buffer (3300 samples * 2 bytes = 6600)
    static const int kSamples = 3300;
    char eciBuffer[3300 * 2] = {};

    // Audio format
    WAVEFORMATEX lastFormat = {};
    bool formatValid = false;

    // 2.0 waveOut callback routing
    DWORD callbackType = 0;
    DWORD_PTR callbackTarget = 0;
    DWORD_PTR callbackInstance = 0;

    // Events
    HANDLE doneEvent = nullptr;
    HANDLE stopEvent = nullptr;
    HANDLE cmdEvent = nullptr;
    HANDLE initEvent = nullptr;
    std::atomic<int> initOk{ 0 };

    // Cancel + generations
    std::atomic<uint32_t> cancelToken{ 1 };
    std::atomic<uint32_t> genCounter{ 1 };
    std::atomic<uint32_t> activeGen{ 0 };
    std::atomic<uint32_t> currentGen{ 0 };

    // Output pacing (2.0 only)
    std::atomic<uint64_t> bytesPerSec{ 0 };

    // Silence trimming (2.0 only): cap consecutive silence to maxSilenceSamples.
    uint32_t silenceSamples = 0;
    uint32_t maxSilenceSamples = 0; // 0 = disabled; set from sample rate after format detection

    // Sonic rate boost: speed up audio without pitch change.
    sonicStream sonicStream = nullptr;
    float rateBoost = 1.0f; // 1.0 = normal, 1.5 = 50% faster, 2.0 = 2x

    // Voice settings (dirty-tracked, applied on worker before synthesis)
    SettingInt vparams[8]; // index 1-7 maps to ECI voice param IDs
    SettingInt variant;
    SettingInt voice; // param 9 (3.3: language ID)

    int currentVariant = 0;
    int currentVoice = 0;

    // Command queue
    std::mutex cmdMtx;
    std::deque<Cmd> cmdQ;
    std::thread worker;

    // Output queue
    std::mutex outMtx;
    std::deque<StreamItem> outQ;
    size_t queuedAudioBytes = 0;
    size_t maxBufferedBytes = 4 * 1024 * 1024;
    size_t maxQueueItems = 8192;
};

static ELOQ_STATE* g_state = nullptr;
static std::mutex g_globalMtx;

// ------------------------------------------------------------
// Helpers
// ------------------------------------------------------------
static bool isCallerFromEloquence(ELOQ_STATE* s, void* ra) {
    if (!s) return false;
    HMODULE caller = nullptr;
    if (!GetModuleHandleExW(
        GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS | GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT,
        reinterpret_cast<LPCWSTR>(ra),
        &caller
    )) {
        dbg("isCallerFromEloquence: GetModuleHandleExW failed for ra=%p", ra);
        return false;
    }
    if (!caller) return false;
    if (caller == s->eciModule) return true;
    if (s->engsynModule && caller == s->engsynModule) return true;
    if (s->cwlModule && caller == s->cwlModule) return true;

    // Log the unknown caller module for debugging.
    char modName[MAX_PATH] = {};
    GetModuleFileNameA(caller, modName, MAX_PATH);
    dbg("isCallerFromEloquence: REJECTED caller=%p (%s) ra=%p", caller, modName, ra);
    return false;
}

static void signalWaveOutMessage(ELOQ_STATE* s, UINT msg, WAVEHDR* hdr) {
    if (!s) return;
    const DWORD cbType = (s->callbackType & CALLBACK_TYPEMASK);

    if (cbType == CALLBACK_FUNCTION) {
        typedef void(CALLBACK* WaveOutProc)(HWAVEOUT, UINT, DWORD_PTR, DWORD_PTR, DWORD_PTR);
        WaveOutProc proc = reinterpret_cast<WaveOutProc>(s->callbackTarget);
        if (proc) {
            proc(reinterpret_cast<HWAVEOUT>(s), msg, s->callbackInstance,
                reinterpret_cast<DWORD_PTR>(hdr), 0);
        }
        return;
    }

    if (cbType == CALLBACK_WINDOW) {
        HWND hwnd = reinterpret_cast<HWND>(s->callbackTarget);
        if (!hwnd) return;
        UINT mmMsg = 0;
        if (msg == WOM_OPEN) mmMsg = MM_WOM_OPEN;
        else if (msg == WOM_CLOSE) mmMsg = MM_WOM_CLOSE;
        else if (msg == WOM_DONE) mmMsg = MM_WOM_DONE;
        if (mmMsg) PostMessageW(hwnd, mmMsg, reinterpret_cast<WPARAM>(s), reinterpret_cast<LPARAM>(hdr));
        return;
    }

    if (cbType == CALLBACK_THREAD) {
        DWORD tid = static_cast<DWORD>(s->callbackTarget);
        UINT mmMsg = 0;
        if (msg == WOM_OPEN) mmMsg = MM_WOM_OPEN;
        else if (msg == WOM_CLOSE) mmMsg = MM_WOM_CLOSE;
        else if (msg == WOM_DONE) mmMsg = MM_WOM_DONE;
        if (mmMsg && tid) PostThreadMessageW(tid, mmMsg, reinterpret_cast<WPARAM>(s), reinterpret_cast<LPARAM>(hdr));
        return;
    }

    if (cbType == CALLBACK_EVENT) {
        HANDLE ev = reinterpret_cast<HANDLE>(s->callbackTarget);
        if (ev) SetEvent(ev);
        return;
    }
}

static void clearOutputQueueLocked(ELOQ_STATE* s) {
    s->outQ.clear();
    s->queuedAudioBytes = 0;
}

static void pushAudioToQueue(ELOQ_STATE* s, uint32_t gen, std::vector<uint8_t> buf);

static void enqueueAudioFromHook(ELOQ_STATE* s, uint32_t gen, const void* data, size_t size) {
    if (!s || !data || size == 0) return;

    const uint8_t* src = static_cast<const uint8_t*>(data);
    std::vector<uint8_t> buf;

    // Silence trimming for mode 20: cap runs of silence to maxSilenceSamples.
    if (s->mode == ELOQ_MODE_20 && s->maxSilenceSamples > 0 && s->formatValid) {
        const int bps = s->lastFormat.wBitsPerSample;
        const int nch = s->lastFormat.nChannels;
        const int frameSize = (bps / 8) * nch;

        if (frameSize > 0 && (bps == 8 || bps == 16)) {
            buf.reserve(size);
            for (size_t i = 0; i + (size_t)frameSize <= size; i += frameSize) {
                bool silent = true;
                if (bps == 8) {
                    for (int c = 0; c < nch && silent; c++) {
                        uint8_t v = src[i + c];
                        if (v < 124 || v > 132) silent = false;
                    }
                } else {
                    for (int c = 0; c < nch && silent; c++) {
                        int16_t v;
                        memcpy(&v, &src[i + c * 2], sizeof(v));
                        if (v < -128 || v > 128) silent = false;
                    }
                }
                if (silent) {
                    if (++s->silenceSamples <= s->maxSilenceSamples)
                        buf.insert(buf.end(), src + i, src + i + frameSize);
                } else {
                    s->silenceSamples = 0;
                    buf.insert(buf.end(), src + i, src + i + frameSize);
                }
            }
            if (buf.empty()) return;
        } else {
            buf.assign(src, src + size);
        }
    } else {
        buf.assign(src, src + size);
    }

    // Sonic rate boost: time-stretch without pitch change.
    if (s->rateBoost > 1.001f && s->formatValid) {
        const int bps = s->lastFormat.wBitsPerSample;
        const int nch = s->lastFormat.nChannels;
        const int frameSize = (bps / 8) * nch;
        if (frameSize > 0 && (bps == 8 || bps == 16)) {
            if (!s->sonicStream) {
                s->sonicStream = sonicCreateStream(s->lastFormat.nSamplesPerSec, nch);
                sonicSetSpeed(s->sonicStream, s->rateBoost);
            }
            int numSamples = (int)(buf.size() / frameSize);
            if (bps == 8)
                sonicWriteUnsignedCharToStream(s->sonicStream, buf.data(), numSamples);
            else
                sonicWriteShortToStream(s->sonicStream, reinterpret_cast<short*>(buf.data()), numSamples);

            int avail = sonicSamplesAvailable(s->sonicStream);
            if (avail > 0) {
                buf.resize(avail * frameSize);
                if (bps == 8)
                    sonicReadUnsignedCharFromStream(s->sonicStream, buf.data(), avail);
                else
                    sonicReadShortFromStream(s->sonicStream, reinterpret_cast<short*>(buf.data()), avail);
            } else {
                return; // Sonic is buffering internally, no output yet.
            }
        }
        if (buf.empty()) return;
    }

    pushAudioToQueue(s, gen, std::move(buf));
}

static void pushAudioToQueue(ELOQ_STATE* s, uint32_t gen, std::vector<uint8_t> buf) {
    if (buf.empty()) return;

    std::lock_guard<std::mutex> g(s->outMtx);
    const uint32_t curGen = s->currentGen.load(std::memory_order_relaxed);
    if (curGen == 0 || gen != curGen) return;

    const size_t limit = s->maxBufferedBytes;
    // Drop oldest audio if full.
    while ((s->queuedAudioBytes + buf.size() > limit) || (s->outQ.size() >= s->maxQueueItems)) {
        bool found = false;
        for (auto it = s->outQ.begin(); it != s->outQ.end(); ++it) {
            if (it->type == ELOQ_ITEM_AUDIO) {
                size_t remaining = (it->data.size() > it->offset) ? (it->data.size() - it->offset) : 0;
                if (s->queuedAudioBytes >= remaining) s->queuedAudioBytes -= remaining;
                else s->queuedAudioBytes = 0;
                s->outQ.erase(it);
                found = true;
                break;
            }
        }
        if (!found) return;
    }

    const size_t bufSize = buf.size();
    StreamItem it(ELOQ_ITEM_AUDIO, 0, gen);
    it.data = std::move(buf);
    s->queuedAudioBytes += bufSize;
    s->outQ.push_back(std::move(it));
}

static void pushMarker(ELOQ_STATE* s, int type, int value, uint32_t gen) {
    std::lock_guard<std::mutex> g(s->outMtx);
    const uint32_t curGen = s->currentGen.load(std::memory_order_relaxed);
    if (curGen == 0 || gen != curGen) return;
    s->outQ.push_back(StreamItem(type, value, gen));
}

// ------------------------------------------------------------
// ECI callback (shared by both modes, called on worker thread)
// ------------------------------------------------------------
static int __cdecl eciCallback(int h, int msgType, int length, void* data) {
    ELOQ_STATE* s = g_state;
    if (!s) { dbg("eciCallback: g_state NULL"); return 2; }

    const uint32_t gen = s->activeGen.load(std::memory_order_relaxed);
    const uint32_t curGen = s->currentGen.load(std::memory_order_relaxed);
    dbg("eciCallback: msg=%d len=%d gen=%u curGen=%u", msgType, length, gen, curGen);
    if (gen == 0 || gen != curGen) { dbg("eciCallback: gen mismatch, dropping"); return 2; }

    if (s->mode == ELOQ_MODE_33 && msgType == 0) {
        if (length > 0) {
            // 3.3: Audio data in output buffer.
            size_t bytes = (size_t)length * 2;
            if (bytes > sizeof(s->eciBuffer)) bytes = sizeof(s->eciBuffer);
            dbg("eciCallback: enqueueing %zu audio bytes", bytes);
            enqueueAudioFromHook(s, gen, s->eciBuffer, bytes);
        } else {
            // 3.3: length==0 means end of synthesis (no 0xFFFF in this mode).
            dbg("eciCallback: DONE (msg=0, len=0)");
            if (s->doneEvent) SetEvent(s->doneEvent);
        }
    }

    if (msgType == 2) {
        if (length == 0xFFFF) {
            // End of utterance (2.0 style).
            dbg("eciCallback: DONE (0xFFFF)");
            if (s->doneEvent) SetEvent(s->doneEvent);
        } else {
            // Index marker.
            dbg("eciCallback: INDEX %d", length);
            pushMarker(s, ELOQ_ITEM_INDEX, length, gen);
        }
    }

    return 1;
}

// ------------------------------------------------------------
// WaveOut hooks (2.0 only)
// ------------------------------------------------------------
static MMRESULT WINAPI hook_waveOutOpen(
    LPHWAVEOUT phwo,
    UINT uDeviceID,
    LPCWAVEFORMATEX pwfx,
    DWORD_PTR dwCallback,
    DWORD_PTR dwInstance,
    DWORD fdwOpen
) {
    ELOQ_STATE* s = g_state;
    // For mode 20: intercept ALL waveOutOpen calls. In the bridge host
    // process, the only waveOut caller is the Eloquence engine (via
    // Speech.dll). NVDA 2026.1's nvwave uses WASAPI, not waveOut.
    bool fromEloq = s && s->mode == ELOQ_MODE_20;
    dbg("hook_waveOutOpen: fromEloq=%d mode=%d", fromEloq, s ? s->mode : -1);
    if (!fromEloq) {
        return g_waveOutOpenOrig ? g_waveOutOpenOrig(phwo, uDeviceID, pwfx, dwCallback, dwInstance, fdwOpen)
            : MMSYSERR_ERROR;
    }

    HWAVEOUT fakeHandle = reinterpret_cast<HWAVEOUT>(s);
    if (phwo) *phwo = fakeHandle;
    s->eloqWaveHandle = fakeHandle;

    if (pwfx) {
        s->lastFormat = *pwfx;
        s->formatValid = true;

        uint64_t bps = (uint64_t)pwfx->nAvgBytesPerSec;
        if (!bps && pwfx->nSamplesPerSec && pwfx->nBlockAlign)
            bps = (uint64_t)pwfx->nSamplesPerSec * (uint64_t)pwfx->nBlockAlign;
        if (!bps) bps = 22050;
        s->bytesPerSec.store(bps, std::memory_order_relaxed);

        // Silence trimming: cap at 60 ms of silence.
        if (s->maxSilenceSamples == 0 && pwfx->nSamplesPerSec > 0)
            s->maxSilenceSamples = pwfx->nSamplesPerSec * 60 / 1000;
    }

    s->callbackType = fdwOpen;
    s->callbackTarget = dwCallback;
    s->callbackInstance = dwInstance;

    signalWaveOutMessage(s, WOM_OPEN, nullptr);
    return MMSYSERR_NOERROR;
}

static MMRESULT WINAPI hook_waveOutPrepareHeader(HWAVEOUT hwo, LPWAVEHDR pwh, UINT cbwh) {
    ELOQ_STATE* s = g_state;
    if (!s || s->mode != ELOQ_MODE_20 || hwo != s->eloqWaveHandle) {
        return g_waveOutPrepareHeaderOrig ? g_waveOutPrepareHeaderOrig(hwo, pwh, cbwh) : MMSYSERR_ERROR;
    }
    if (pwh) pwh->dwFlags |= WHDR_PREPARED;
    return MMSYSERR_NOERROR;
}

static MMRESULT WINAPI hook_waveOutUnprepareHeader(HWAVEOUT hwo, LPWAVEHDR pwh, UINT cbwh) {
    ELOQ_STATE* s = g_state;
    if (!s || s->mode != ELOQ_MODE_20 || hwo != s->eloqWaveHandle) {
        return g_waveOutUnprepareHeaderOrig ? g_waveOutUnprepareHeaderOrig(hwo, pwh, cbwh) : MMSYSERR_ERROR;
    }
    if (pwh) pwh->dwFlags &= ~WHDR_PREPARED;
    return MMSYSERR_NOERROR;
}

static MMRESULT WINAPI hook_waveOutWrite(HWAVEOUT hwo, LPWAVEHDR pwh, UINT cbwh) {
    ELOQ_STATE* s = g_state;
    bool fromEloq = s && s->mode == ELOQ_MODE_20 && hwo == s->eloqWaveHandle;
    if (!fromEloq) {
        return g_waveOutWriteOrig ? g_waveOutWriteOrig(hwo, pwh, cbwh) : MMSYSERR_ERROR;
    }

    if (!pwh) return MMSYSERR_INVALPARAM;

    const uint32_t gen = s->activeGen.load(std::memory_order_relaxed);
    const uint32_t curGen = s->currentGen.load(std::memory_order_relaxed);
    const bool capturing = (gen != 0 && gen == curGen);

    dbg("hook_waveOutWrite: %lu bytes, capturing=%d gen=%u curGen=%u",
        pwh->dwBufferLength, capturing, gen, curGen);

    if (capturing && pwh->lpData && pwh->dwBufferLength > 0) {
        enqueueAudioFromHook(s, gen, pwh->lpData, (size_t)pwh->dwBufferLength);
    }

    pwh->dwFlags |= WHDR_DONE;
    signalWaveOutMessage(s, WOM_DONE, pwh);
    return MMSYSERR_NOERROR;
}

static MMRESULT WINAPI hook_waveOutReset(HWAVEOUT hwo) {
    ELOQ_STATE* s = g_state;
    if (!s || s->mode != ELOQ_MODE_20 || hwo != s->eloqWaveHandle) {
        return g_waveOutResetOrig ? g_waveOutResetOrig(hwo) : MMSYSERR_ERROR;
    }
    dbg("hook_waveOutReset: signaling doneEvent");
    // Engine finished playback — signal done.
    if (s->doneEvent) SetEvent(s->doneEvent);
    return MMSYSERR_NOERROR;
}

static MMRESULT WINAPI hook_waveOutClose(HWAVEOUT hwo) {
    ELOQ_STATE* s = g_state;
    if (!s || s->mode != ELOQ_MODE_20 || hwo != s->eloqWaveHandle) {
        return g_waveOutCloseOrig ? g_waveOutCloseOrig(hwo) : MMSYSERR_ERROR;
    }
    dbg("hook_waveOutClose: signaling doneEvent");
    // Engine closing waveOut device — signal done as fallback.
    if (s->doneEvent) SetEvent(s->doneEvent);
    signalWaveOutMessage(s, WOM_CLOSE, nullptr);
    return MMSYSERR_NOERROR;
}

static bool ensureHooksInstalled() {
    static LONG once = 0;
    if (InterlockedCompareExchange(&once, 1, 0) != 0) return true;

    LoadLibraryW(L"winmm.dll");
    LoadLibraryW(L"winmmbase.dll");

    MH_STATUS st = MH_Initialize();
    if (st != MH_OK && st != MH_ERROR_ALREADY_INITIALIZED) return false;

    auto tryHookApi = [&](const wchar_t* mod, const char* proc, LPVOID detour, LPVOID* orig) -> bool {
        MH_STATUS rc = MH_CreateHookApi(mod, proc, detour, orig);
        return (rc == MH_OK || rc == MH_ERROR_ALREADY_CREATED);
    };

    auto hookEither = [&](const char* proc, LPVOID detour, LPVOID* orig) -> bool {
        if (tryHookApi(L"winmm.dll", proc, detour, orig)) return true;
        if (tryHookApi(L"winmmbase.dll", proc, detour, orig)) return true;
        return false;
    };

    const bool okOpen   = hookEither("waveOutOpen",             (LPVOID)hook_waveOutOpen,             (LPVOID*)&g_waveOutOpenOrig);
    const bool okPrep   = hookEither("waveOutPrepareHeader",    (LPVOID)hook_waveOutPrepareHeader,    (LPVOID*)&g_waveOutPrepareHeaderOrig);
    const bool okUnprep = hookEither("waveOutUnprepareHeader",  (LPVOID)hook_waveOutUnprepareHeader,  (LPVOID*)&g_waveOutUnprepareHeaderOrig);
    const bool okWrite  = hookEither("waveOutWrite",            (LPVOID)hook_waveOutWrite,            (LPVOID*)&g_waveOutWriteOrig);
    const bool okReset  = hookEither("waveOutReset",            (LPVOID)hook_waveOutReset,            (LPVOID*)&g_waveOutResetOrig);
    const bool okClose  = hookEither("waveOutClose",            (LPVOID)hook_waveOutClose,            (LPVOID*)&g_waveOutCloseOrig);

    if (!(okOpen && okPrep && okUnprep && okWrite && okReset && okClose)) {
        return false;
    }

    MH_EnableHook(MH_ALL_HOOKS);
    return true;
}

// ------------------------------------------------------------
// ELOQ.CFG path patching (3.3 only)
// ------------------------------------------------------------
static void patchEloqCfg(const std::wstring& dir) {
    std::wstring cfgPath = dir + L"\\ELOQ.CFG";
    HANDLE hFile = CreateFileW(cfgPath.c_str(), GENERIC_READ | GENERIC_WRITE,
        FILE_SHARE_READ, nullptr, OPEN_EXISTING, 0, nullptr);
    if (hFile == INVALID_HANDLE_VALUE) return;

    DWORD fileSize = GetFileSize(hFile, nullptr);
    if (fileSize == INVALID_FILE_SIZE || fileSize < 2200) {
        CloseHandle(hFile);
        return;
    }

    std::vector<char> buf(fileSize);
    DWORD bytesRead = 0;
    if (!ReadFile(hFile, buf.data(), fileSize, &bytesRead, nullptr) || bytesRead != fileSize) {
        CloseHandle(hFile);
        return;
    }

    // Convert dir to MBCS for comparison/replacement.
    char mbcsDir[MAX_PATH * 2] = {};
    WideCharToMultiByte(CP_ACP, 0, dir.c_str(), -1, mbcsDir, sizeof(mbcsDir), nullptr, nullptr);
    size_t mbcsDirLen = strlen(mbcsDir);
    // Ensure trailing backslash.
    if (mbcsDirLen > 0 && mbcsDir[mbcsDirLen - 1] != '\\') {
        mbcsDir[mbcsDirLen] = '\\';
        mbcsDir[mbcsDirLen + 1] = '\0';
        mbcsDirLen++;
    }

    // Check offset 2119 for a path reference and patch if needed.
    const size_t offset = 2119;
    if (offset >= fileSize) {
        CloseHandle(hFile);
        return;
    }

    // Read the line at offset 2119.
    size_t lineEnd = offset;
    while (lineEnd < fileSize && buf[lineEnd] != '\n' && buf[lineEnd] != '\r') lineEnd++;
    std::string currentLine(buf.data() + offset, lineEnd - offset);

    // Check if line already starts with the correct path.
    if (currentLine.size() >= mbcsDirLen &&
        _strnicmp(currentLine.c_str(), mbcsDir, mbcsDirLen) == 0) {
        CloseHandle(hFile);
        return;
    }

    // Find the old path prefix (everything before the last backslash-filename).
    size_t lastSlash = currentLine.rfind('\\');
    if (lastSlash == std::string::npos) {
        CloseHandle(hFile);
        return;
    }
    std::string oldPrefix = currentLine.substr(0, lastSlash + 1);

    // Replace all occurrences of oldPrefix with mbcsDir in the entire buffer.
    std::string content(buf.data(), fileSize);
    std::string newPrefix(mbcsDir, mbcsDirLen);
    size_t pos = 0;
    while ((pos = content.find(oldPrefix, pos)) != std::string::npos) {
        content.replace(pos, oldPrefix.size(), newPrefix);
        pos += newPrefix.size();
    }

    // Write back.
    SetFilePointer(hFile, 0, nullptr, FILE_BEGIN);
    DWORD written = 0;
    WriteFile(hFile, content.data(), (DWORD)content.size(), &written, nullptr);
    SetEndOfFile(hFile);
    CloseHandle(hFile);
}

// ------------------------------------------------------------
// Version detection
// ------------------------------------------------------------
static int detectMode(const std::wstring& dir) {
    // 2.0: has ENGSYN32.DLL
    std::wstring engsynPath = dir + L"\\ENGSYN32.DLL";
    if (GetFileAttributesW(engsynPath.c_str()) != INVALID_FILE_ATTRIBUTES)
        return ELOQ_MODE_20;

    // 3.3: has .SYN files
    WIN32_FIND_DATAW fd;
    std::wstring pattern = dir + L"\\*.SYN";
    HANDLE hFind = FindFirstFileW(pattern.c_str(), &fd);
    if (hFind != INVALID_HANDLE_VALUE) {
        FindClose(hFind);
        return ELOQ_MODE_33;
    }

    return ELOQ_MODE_NONE;
}

// ------------------------------------------------------------
// DLL loading and ECI function resolution
// ------------------------------------------------------------
static bool resolveEciFunctions(ELOQ_STATE* s) {
    HMODULE m = s->eciModule;
    if (!m) return false;

    s->fnNew              = (eciNewFunc)GetProcAddress(m, "eciNew");
    s->fnDelete           = (eciDeleteFunc)GetProcAddress(m, "eciDelete");
    s->fnRequestLicense   = (eciRequestLicenseFunc)GetProcAddress(m, "eciRequestLicense");
    s->fnSetOutputBuffer  = (eciSetOutputBufferFunc)GetProcAddress(m, "eciSetOutputBuffer");
    s->fnSetOutputDevice  = (eciSetOutputDeviceFunc)GetProcAddress(m, "eciSetOutputDevice");
    s->fnRegisterCallback = (eciRegisterCallbackFunc)GetProcAddress(m, "eciRegisterCallback");
    s->fnSetParam         = (eciSetParamFunc)GetProcAddress(m, "eciSetParam");
    s->fnGetParam         = (eciGetParamFunc)GetProcAddress(m, "eciGetParam");
    s->fnSetVoiceParam    = (eciSetVoiceParamFunc)GetProcAddress(m, "eciSetVoiceParam");
    s->fnGetVoiceParam    = (eciGetVoiceParamFunc)GetProcAddress(m, "eciGetVoiceParam");
    s->fnCopyVoice        = (eciCopyVoiceFunc)GetProcAddress(m, "eciCopyVoice");
    s->fnAddText          = (eciAddTextFunc)GetProcAddress(m, "eciAddText");
    s->fnInsertIndex      = (eciInsertIndexFunc)GetProcAddress(m, "eciInsertIndex");
    s->fnSynthesize       = (eciSynthesizeFunc)GetProcAddress(m, "eciSynthesize");
    s->fnStop             = (eciStopFunc)GetProcAddress(m, "eciStop");
    s->fnSpeaking         = (eciSpeakingFunc)GetProcAddress(m, "eciSpeaking");
    s->fnSynchronize      = (eciSynchronizeFunc)GetProcAddress(m, "eciSynchronize");
    s->fnVersion          = (eciVersionFunc)GetProcAddress(m, "eciVersion");
    s->fnNewDict          = (eciNewDictFunc)GetProcAddress(m, "eciNewDict");
    s->fnSetDict          = (eciSetDictFunc)GetProcAddress(m, "eciSetDict");
    s->fnLoadDict         = (eciLoadDictFunc)GetProcAddress(m, "eciLoadDict");

    // Minimum required.
    return s->fnNew && s->fnDelete && s->fnSetParam && s->fnAddText &&
           s->fnSynthesize && s->fnStop && s->fnRegisterCallback;
}

static void* tryLicense33(ELOQ_STATE* s) {
    if (!s->fnRequestLicense || !s->fnNew) return nullptr;

    auto tryOffset = [&](int offset) -> void* {
        int lt = (int)time(nullptr) + offset;
        lt ^= 0x39AB43F2;
        s->fnRequestLicense(lt);
        return s->fnNew();
    };

    void* h = tryOffset(0);
    if (!h) h = tryOffset(+3600);
    if (!h) h = tryOffset(-3600);
    if (!h) h = s->fnNew(); // last resort without license
    return h;
}

// ------------------------------------------------------------
// Worker thread: apply settings, synthesize, wait for done
// ------------------------------------------------------------
static void applyDirtySettings(ELOQ_STATE* s) {
    if (!s || !s->handle) return;

    // Voice/language change (3.3 only, param 9).
    if (s->mode == ELOQ_MODE_33 && s->voice.dirty.exchange(0, std::memory_order_relaxed)) {
        int v = s->voice.value.load(std::memory_order_relaxed);
        if (v != s->currentVoice && s->fnSetParam) {
            s->fnSetParam(s->handle, 9, v);
            s->currentVoice = v;
        }
    }

    // Variant change (eciCopyVoice).
    if (s->variant.dirty.exchange(0, std::memory_order_relaxed)) {
        int v = s->variant.value.load(std::memory_order_relaxed);
        if (v != s->currentVariant && s->fnCopyVoice) {
            s->fnCopyVoice(s->handle, v, 0);
            s->currentVariant = v;
        }
    }

    // Voice parameters 1-7.
    for (int i = 1; i <= 7; i++) {
        if (s->vparams[i].dirty.exchange(0, std::memory_order_relaxed)) {
            int v = s->vparams[i].value.load(std::memory_order_relaxed);
            if (s->fnSetVoiceParam)
                s->fnSetVoiceParam(s->handle, 0, i, v);
        }
    }
}

static void workerLoop(ELOQ_STATE* s) {
    if (!s) return;

    s->workerThreadId = GetCurrentThreadId();
    SetThreadPriority(GetCurrentThread(), THREAD_PRIORITY_ABOVE_NORMAL);

    // Initialize Windows message queue (some ECI internals use window messages).
    MSG msg;
    PeekMessageW(&msg, nullptr, 0, 0, PM_NOREMOVE);

    // ---- Load DLLs and create ECI handle ----
    bool ok = false;
    std::wstring eciPath = s->dllDir + L"\\ECI32D.DLL";
    std::wstring cwlPath = s->dllDir + L"\\CW3220MT.DLL";

    // For 2.0: install hooks BEFORE loading DLLs (ENGSYN32 may init early).
    if (s->mode == ELOQ_MODE_20) {
        dbg("worker: installing waveOut hooks for mode 20");
        if (!ensureHooksInstalled()) {
            dbg("worker: hook installation FAILED");
            s->initOk.store(-1, std::memory_order_relaxed);
            if (s->initEvent) SetEvent(s->initEvent);
            return;
        }
        dbg("worker: hooks installed OK");
    }

    // Set DLL search directory so implicit dependencies (e.g. ENGSYN32.DLL
    // imported by ECI32D.DLL in 2.0 mode) are found in the engine folder.
    dbg("worker: SetDllDirectoryW('%ls')", s->dllDir.c_str());
    SetDllDirectoryW(s->dllDir.c_str());

    // Load Borland runtime first (all other DLLs depend on it).
    dbg("worker: loading CW3220MT.DLL...");
    if (GetFileAttributesW(cwlPath.c_str()) != INVALID_FILE_ATTRIBUTES) {
        s->cwlModule = LoadLibraryW(cwlPath.c_str());
        dbg("worker: CW3220MT.DLL = %p", s->cwlModule);
    } else {
        dbg("worker: CW3220MT.DLL not found, skipping");
    }

    // Load ECI32D.DLL.
    dbg("worker: loading ECI32D.DLL...");
    s->eciModule = LoadLibraryW(eciPath.c_str());
    dbg("worker: ECI32D.DLL = %p (err=%lu)", s->eciModule, s->eciModule ? 0 : GetLastError());

    // Restore default DLL search order.
    SetDllDirectoryW(nullptr);
    if (!s->eciModule) {
        s->initOk.store(-1, std::memory_order_relaxed);
        if (s->initEvent) SetEvent(s->initEvent);
        return;
    }

    // For 2.0: get handle to ENGSYN32.DLL (loaded as ECI32D import).
    // Speech.dll is loaded lazily during priming, so we grab it later.
    if (s->mode == ELOQ_MODE_20) {
        s->engsynModule = GetModuleHandleW(L"ENGSYN32.DLL");
        dbg("worker: ENGSYN32.DLL = %p", s->engsynModule);
    }

    dbg("worker: resolving ECI functions...");
    if (!resolveEciFunctions(s)) {
        dbg("worker: resolveEciFunctions FAILED");
        s->initOk.store(-1, std::memory_order_relaxed);
        if (s->initEvent) SetEvent(s->initEvent);
        return;
    }
    dbg("worker: ECI functions resolved OK");

    // Create ECI handle.
    dbg("worker: creating ECI handle (mode=%d)...", s->mode);
    if (s->mode == ELOQ_MODE_33) {
        patchEloqCfg(s->dllDir);
        s->handle = tryLicense33(s);
    } else {
        s->handle = s->fnNew();
    }
    dbg("worker: ECI handle = %p", s->handle);

    if (!s->handle) {
        s->initOk.store(-1, std::memory_order_relaxed);
        if (s->initEvent) SetEvent(s->initEvent);
        return;
    }

    // ---- Mode-specific setup ----
    if (s->mode == ELOQ_MODE_33) {
        // Register callback FIRST (matches old driver order).
        dbg("worker: RegisterCallback(fn=%p)", (void*)eciCallback);
        int cbRc = s->fnRegisterCallback(s->handle, (void*)eciCallback, nullptr);
        dbg("worker: RegisterCallback returned %d", cbRc);

        // Set output buffer for callback audio delivery.
        dbg("worker: 3.3 setup — SetOutputBuffer(%d samples, buf=%p)", s->kSamples, s->eciBuffer);
        if (s->fnSetOutputBuffer) {
            int obRc = s->fnSetOutputBuffer(s->handle, s->kSamples, s->eciBuffer);
            dbg("worker: SetOutputBuffer returned %d", obRc);
        }

        // Synth mode: param 1, value 1 (to-buffer, not to speakers).
        int smRc = s->fnSetParam(s->handle, 1, 1);
        dbg("worker: SetParam(1,1) [synth mode] returned %d", smRc);
        // Verify it was actually set.
        if (s->fnGetParam) {
            int actual = s->fnGetParam(s->handle, 1);
            dbg("worker: GetParam(1) = %d (expect 1 for buffer mode)", actual);
        }

        // Set known format: 11025 Hz, 16-bit, mono.
        memset(&s->lastFormat, 0, sizeof(s->lastFormat));
        s->lastFormat.wFormatTag = WAVE_FORMAT_PCM;
        s->lastFormat.nChannels = 1;
        s->lastFormat.nSamplesPerSec = 11025;
        s->lastFormat.wBitsPerSample = 16;
        s->lastFormat.nBlockAlign = 2;
        s->lastFormat.nAvgBytesPerSec = 22050;
        s->formatValid = true;
        s->bytesPerSec.store(22050, std::memory_order_relaxed);

    } else { // ELOQ_MODE_20
        dbg("worker: 2.0 setup — SetOutputDevice...");
        // Set output device (optional, may fail).
        if (s->fnSetOutputDevice)
            s->fnSetOutputDevice(s->handle, 0);

        dbg("worker: 2.0 SetParam(1,1)...");
        // Synth mode: param 1, value 1.
        s->fnSetParam(s->handle, 1, 1);

        // Prime: add space, synthesize, poll, synchronize, stop.
        // This is required for 2.0 to initialize its internal state.
        dbg("worker: 2.0 priming...");
        s->fnAddText(s->handle, " ");
        dbg("worker: 2.0 fnSynthesize...");
        s->fnSynthesize(s->handle);
        dbg("worker: 2.0 waiting for speaking to finish...");
        if (s->fnSpeaking) {
            while (s->fnSpeaking(s->handle)) {
                Sleep(5);
            }
        }
        dbg("worker: 2.0 speaking done, skipping fnSynchronize (crashes with hooked waveOut)");
        s->fnStop(s->handle);

        // Register callback AFTER priming (2.0 requirement).
        s->fnRegisterCallback(s->handle, (void*)eciCallback, nullptr);
    }

    // Read initial voice params.
    if (s->fnGetVoiceParam) {
        for (int i = 1; i <= 7; i++) {
            int v = s->fnGetVoiceParam(s->handle, 0, i);
            s->vparams[i].value.store(v, std::memory_order_relaxed);
        }
    }
    if (s->mode == ELOQ_MODE_33 && s->fnGetParam) {
        s->currentVoice = s->fnGetParam(s->handle, 9);
        s->voice.value.store(s->currentVoice, std::memory_order_relaxed);
    }

    s->initOk.store(1, std::memory_order_relaxed);
    if (s->initEvent) SetEvent(s->initEvent);

    dbg("worker: init OK, mode=%d handle=%p genCounter=%u cancelToken=%u",
        s->mode, s->handle,
        s->genCounter.load(std::memory_order_relaxed),
        s->cancelToken.load(std::memory_order_relaxed));

    // ---- Main command loop ----
    auto pumpMessages = [&]() {
        while (PeekMessageW(&msg, nullptr, 0, 0, PM_REMOVE)) {
            TranslateMessage(&msg);
            DispatchMessageW(&msg);
        }
    };

    while (true) {
        pumpMessages();

        Cmd cmd;
        bool hasCmd = false;
        {
            std::lock_guard<std::mutex> lk(s->cmdMtx);
            if (!s->cmdQ.empty()) {
                cmd = std::move(s->cmdQ.front());
                s->cmdQ.pop_front();
                hasCmd = true;
            } else {
                ResetEvent(s->cmdEvent);
            }
        }

        if (!hasCmd) {
            MsgWaitForMultipleObjectsEx(
                1,
                &s->cmdEvent,
                INFINITE,
                QS_ALLINPUT,
                MWMO_INPUTAVAILABLE
            );
            continue;
        }

        if (cmd.type == Cmd::CMD_QUIT) {
            dbg("worker: CMD_QUIT");
            break;
        }

        // Check if this command was canceled before we process it.
        const uint32_t snap = s->cancelToken.load(std::memory_order_relaxed);
        dbg("worker: CMD_SPEAK snap=%u cmdSnap=%u text='%.80s'", snap, cmd.cancelSnapshot, cmd.text.c_str());
        if (cmd.cancelSnapshot != snap) {
            dbg("worker: command canceled (snap mismatch)");
            continue;
        }

        const uint32_t gen = s->genCounter.fetch_add(1, std::memory_order_relaxed);
        dbg("worker: gen=%u", gen);

        ResetEvent(s->stopEvent);
        ResetEvent(s->doneEvent);
        s->silenceSamples = 0;

        // Gate on.
        s->currentGen.store(gen, std::memory_order_relaxed);
        s->activeGen.store(gen, std::memory_order_relaxed);

        // Clear output queue.
        {
            std::lock_guard<std::mutex> g(s->outMtx);
            clearOutputQueueLocked(s);
        }

        // Apply pending settings.
        applyDirtySettings(s);

        // Send text.
        if (cmd.text.empty()) {
            dbg("worker: empty text, pushing DONE");
            s->activeGen.store(0, std::memory_order_relaxed);
            pushMarker(s, ELOQ_ITEM_DONE, 0, gen);
            continue;
        }

        // Strip brackets/parens for all modes — Eloquence reads them as
        // full words (e.g. "LEFT PAREN LEFT PARENTHESIS").
        // Backtick only stripped for mode 20; mode 33 uses it as ECI
        // inline command prefix (e.g. `da0, `vv92).
        for (auto& ch : cmd.text) {
            switch (ch) {
                case '(': case ')':
                case '{': case '}':
                case '[': case ']':
                    ch = ' ';
                    break;
                case '`':
                    if (s->mode == ELOQ_MODE_20) ch = ' ';
                    break;
            }
        }

        dbg("worker: fnAddText...");
        int addRc = s->fnAddText(s->handle, cmd.text.c_str());
        dbg("worker: fnAddText returned %d", addRc);
        dbg("worker: fnSynthesize...");
        int synRc = s->fnSynthesize(s->handle);
        dbg("worker: fnSynthesize returned %d", synRc);

        // Check output queue size after synthesis (for 3.3 synchronous mode).
        {
            std::lock_guard<std::mutex> g(s->outMtx);
            dbg("worker: after synth outQ.size=%zu queuedBytes=%zu currentGen=%u",
                s->outQ.size(), s->queuedAudioBytes,
                s->currentGen.load(std::memory_order_relaxed));
        }

        // Wait for synthesis to complete, pumping messages.
        // - 3.3: ECI callback delivers done via msg queue → doneEvent.
        // - 2.0: doneEvent set by hook_waveOutReset when engine finishes.
        HANDLE waits[2] = { s->doneEvent, s->stopEvent };
        bool stopped = false;

        const DWORD timeout = 120000; // 2 minutes max per utterance.
        DWORD deadline = GetTickCount() + timeout;
        bool waitDone = false;
        while (!waitDone) {
            DWORD remaining = deadline - GetTickCount();
            if ((int)remaining <= 0) {
                dbg("worker: TIMEOUT waiting for synthesis");
                stopped = true;
                break;
            }
            DWORD w = MsgWaitForMultipleObjectsEx(
                2, waits, remaining, QS_ALLINPUT, MWMO_INPUTAVAILABLE);
            if (w == WAIT_OBJECT_0) {
                dbg("worker: doneEvent signaled");
                waitDone = true;
            } else if (w == WAIT_OBJECT_0 + 1) {
                dbg("worker: stopEvent signaled");
                stopped = true;
                waitDone = true;
            } else if (w == WAIT_OBJECT_0 + 2) {
                // Messages available — pump them.
                pumpMessages();
            } else {
                dbg("worker: wait returned %lu", w);
                stopped = true;
                waitDone = true;
            }
        }

        if (stopped || s->cancelToken.load(std::memory_order_relaxed) != snap) {
            dbg("worker: calling fnStop (stopped=%d)", stopped);
            if (s->fnStop) s->fnStop(s->handle);
        }

        // Flush sonic stream to get any remaining buffered audio.
        if (s->sonicStream && s->rateBoost > 1.001f && s->formatValid) {
            sonicFlushStream(s->sonicStream);
            const int bps = s->lastFormat.wBitsPerSample;
            const int nch = s->lastFormat.nChannels;
            const int frameSize = (bps / 8) * nch;
            int avail = sonicSamplesAvailable(s->sonicStream);
            if (avail > 0 && frameSize > 0) {
                std::vector<uint8_t> tail(avail * frameSize);
                if (bps == 8)
                    sonicReadUnsignedCharFromStream(s->sonicStream, tail.data(), avail);
                else
                    sonicReadShortFromStream(s->sonicStream, reinterpret_cast<short*>(tail.data()), avail);
                pushAudioToQueue(s, gen, std::move(tail));
            }
        }

        s->activeGen.store(0, std::memory_order_relaxed);
        pushMarker(s, ELOQ_ITEM_DONE, 0, gen);
        dbg("worker: pushed DONE marker, currentGen=%u", s->currentGen.load(std::memory_order_relaxed));
    }

    // Cleanup.
    if (s->handle) {
        if (s->fnStop) s->fnStop(s->handle);
        if (s->fnDelete) s->fnDelete(s->handle);
        s->handle = nullptr;
    }
}

// ============================================================
// Public API
// ============================================================
extern "C" ELOQ_API int __cdecl eloq_init(const wchar_t* dllDir) {
    if (!dllDir) return -1;

    dbg("eloq_init called");

    std::lock_guard<std::mutex> glk(g_globalMtx);
    if (g_state) { dbg("eloq_init: already initialized"); return 0; }

    int mode = detectMode(dllDir);
    dbg("eloq_init: mode=%d", mode);
    if (mode == ELOQ_MODE_NONE) return -2;

    ELOQ_STATE* s = new ELOQ_STATE();
    s->mode = mode;
    s->dllDir = dllDir;

    s->doneEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    s->stopEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    s->cmdEvent  = CreateEventW(nullptr, TRUE, FALSE, nullptr);
    s->initEvent = CreateEventW(nullptr, TRUE, FALSE, nullptr);

    g_state = s;

    s->worker = std::thread(workerLoop, s);

    // Wait for init.
    WaitForSingleObject(s->initEvent, 10000);

    if (s->initOk.load(std::memory_order_relaxed) != 1) {
        // Init failed. Clean up.
        if (s->worker.joinable()) {
            {
                std::lock_guard<std::mutex> lk(s->cmdMtx);
                Cmd quit;
                quit.type = Cmd::CMD_QUIT;
                s->cmdQ.push_back(quit);
                SetEvent(s->cmdEvent);
            }
            s->worker.join();
        }
        if (s->doneEvent) CloseHandle(s->doneEvent);
        if (s->stopEvent) CloseHandle(s->stopEvent);
        if (s->cmdEvent) CloseHandle(s->cmdEvent);
        if (s->initEvent) CloseHandle(s->initEvent);
        delete s;
        g_state = nullptr;
        return -3;
    }

    return 0;
}

extern "C" ELOQ_API void __cdecl eloq_free(void) {
    std::lock_guard<std::mutex> glk(g_globalMtx);
    ELOQ_STATE* s = g_state;
    if (!s) return;

    // Send quit command.
    {
        std::lock_guard<std::mutex> lk(s->cmdMtx);
        Cmd quit;
        quit.type = Cmd::CMD_QUIT;
        s->cmdQ.push_back(quit);
        SetEvent(s->cmdEvent);
    }

    if (s->worker.joinable()) {
        s->worker.join();
    }

    if (s->sonicStream) {
        sonicDestroyStream(s->sonicStream);
        s->sonicStream = nullptr;
    }

    if (s->doneEvent) CloseHandle(s->doneEvent);
    if (s->stopEvent) CloseHandle(s->stopEvent);
    if (s->cmdEvent) CloseHandle(s->cmdEvent);
    if (s->initEvent) CloseHandle(s->initEvent);

    g_state = nullptr;
    delete s;
}

extern "C" ELOQ_API int __cdecl eloq_version(void) {
    ELOQ_STATE* s = g_state;
    return s ? s->mode : 0;
}

extern "C" ELOQ_API int __cdecl eloq_format(int* rate, int* bits, int* channels) {
    ELOQ_STATE* s = g_state;
    if (!s || !s->formatValid) return -1;

    if (rate) *rate = (int)s->lastFormat.nSamplesPerSec;
    if (bits) *bits = (int)s->lastFormat.wBitsPerSample;
    if (channels) *channels = (int)s->lastFormat.nChannels;
    return 0;
}

extern "C" ELOQ_API int __cdecl eloq_speak(const char* text) {
    ELOQ_STATE* s = g_state;
    if (!s || !text) return -1;

    dbg("eloq_speak: '%.80s'", text);

    // Cancel any previous utterance.
    uint32_t newCancel = s->cancelToken.fetch_add(1, std::memory_order_relaxed) + 1;
    SetEvent(s->stopEvent);

    // Enqueue.
    Cmd cmd;
    cmd.type = Cmd::CMD_SPEAK;
    cmd.cancelSnapshot = s->cancelToken.load(std::memory_order_relaxed);
    cmd.text = text;

    dbg("eloq_speak: cancelToken=%u cmdSnap=%u", newCancel, cmd.cancelSnapshot);

    {
        std::lock_guard<std::mutex> lk(s->cmdMtx);
        s->cmdQ.push_back(std::move(cmd));
        SetEvent(s->cmdEvent);
    }
    return 0;
}

extern "C" ELOQ_API int __cdecl eloq_stop(void) {
    ELOQ_STATE* s = g_state;
    if (!s) return -1;

    s->cancelToken.fetch_add(1, std::memory_order_relaxed);
    SetEvent(s->stopEvent);

    // Clear command queue.
    {
        std::lock_guard<std::mutex> lk(s->cmdMtx);
        s->cmdQ.clear();
    }

    // Clear output queue.
    {
        std::lock_guard<std::mutex> lk(s->outMtx);
        clearOutputQueueLocked(s);
    }

    s->currentGen.store(0, std::memory_order_relaxed);
    s->activeGen.store(0, std::memory_order_relaxed);

    return 0;
}

extern "C" ELOQ_API int __cdecl eloq_read(void* buf, int maxBytes, int* itemType, int* value) {
    if (itemType) *itemType = ELOQ_ITEM_NONE;
    if (value) *value = 0;

    ELOQ_STATE* s = g_state;
    if (!s || !buf || maxBytes < 0) return 0;

    std::lock_guard<std::mutex> g(s->outMtx);

    const uint32_t curGen = s->currentGen.load(std::memory_order_relaxed);
    if (curGen == 0) {
        static int readZeroCount = 0;
        if (readZeroCount++ < 5) dbg("eloq_read: currentGen=0, returning NONE");
        clearOutputQueueLocked(s);
        return 0;
    }

    // Drop stale items.
    while (!s->outQ.empty() && s->outQ.front().gen != curGen) {
        StreamItem& it = s->outQ.front();
        if (it.type == ELOQ_ITEM_AUDIO) {
            size_t remaining = (it.data.size() > it.offset) ? (it.data.size() - it.offset) : 0;
            if (s->queuedAudioBytes >= remaining) s->queuedAudioBytes -= remaining;
            else s->queuedAudioBytes = 0;
        }
        s->outQ.pop_front();
    }

    if (s->outQ.empty()) return 0;

    StreamItem& front = s->outQ.front();
    if (itemType) *itemType = front.type;
    if (value) *value = front.value;

    if (front.type == ELOQ_ITEM_AUDIO) {
        size_t remainingSz = (front.data.size() > front.offset) ? (front.data.size() - front.offset) : 0;
        int n = (remainingSz > (size_t)maxBytes) ? maxBytes : (int)remainingSz;

        if (n > 0) {
            std::memcpy(buf, front.data.data() + front.offset, (size_t)n);
            front.offset += (size_t)n;
            if (s->queuedAudioBytes >= (size_t)n) s->queuedAudioBytes -= (size_t)n;
            else s->queuedAudioBytes = 0;
        }

        if (front.offset >= front.data.size()) {
            s->outQ.pop_front();
        }
        return n;
    }

    // DONE / INDEX / ERROR: consume and return.
    s->outQ.pop_front();
    return 0;
}

extern "C" ELOQ_API int __cdecl eloq_set_variant(int variant) {
    ELOQ_STATE* s = g_state;
    if (!s) return -1;
    s->variant.value.store(variant, std::memory_order_relaxed);
    s->variant.dirty.store(1, std::memory_order_relaxed);
    return 0;
}

extern "C" ELOQ_API int __cdecl eloq_set_vparam(int param, int val) {
    ELOQ_STATE* s = g_state;
    if (!s || param < 1 || param > 7) return -1;
    s->vparams[param].value.store(val, std::memory_order_relaxed);
    s->vparams[param].dirty.store(1, std::memory_order_relaxed);
    return 0;
}

extern "C" ELOQ_API int __cdecl eloq_get_vparam(int param) {
    ELOQ_STATE* s = g_state;
    if (!s || param < 1 || param > 7) return -1;
    return s->vparams[param].value.load(std::memory_order_relaxed);
}

extern "C" ELOQ_API int __cdecl eloq_set_voice(int voiceId) {
    ELOQ_STATE* s = g_state;
    if (!s) return -1;
    if (s->mode != ELOQ_MODE_33) return 0; // no-op on 2.0
    s->voice.value.store(voiceId, std::memory_order_relaxed);
    s->voice.dirty.store(1, std::memory_order_relaxed);
    return 0;
}

extern "C" ELOQ_API int __cdecl eloq_set_rate_boost(int percent) {
    ELOQ_STATE* s = g_state;
    if (!s) return -1;
    if (percent < 100) percent = 100;
    if (percent > 600) percent = 600;
    float newRate = (float)percent / 100.0f;
    s->rateBoost = newRate;
    // Update sonic stream speed if it exists.
    if (s->sonicStream) {
        if (newRate > 1.001f) {
            sonicSetSpeed(s->sonicStream, newRate);
        } else {
            // Boost disabled — destroy sonic stream.
            sonicDestroyStream(s->sonicStream);
            s->sonicStream = nullptr;
        }
    }
    dbg("eloq_set_rate_boost: %d%% (%.2f)", percent, newRate);
    return 0;
}

extern "C" ELOQ_API int __cdecl eloq_get_rate_boost() {
    ELOQ_STATE* s = g_state;
    if (!s) return 100;
    return (int)(s->rateBoost * 100.0f + 0.5f);
}

extern "C" ELOQ_API int __cdecl eloq_load_dict(const char* mainPath, const char* rootPath) {
    ELOQ_STATE* s = g_state;
    if (!s || s->mode != ELOQ_MODE_33) return -1;
    if (!s->fnNewDict || !s->fnSetDict || !s->fnLoadDict) return -1;
    // Dict loading happens on any thread (ECI is tolerant for this).
    // But we only do it once at init or on explicit call.
    if (s->dictHandle < 0 && s->handle) {
        s->dictHandle = s->fnNewDict(s->handle);
        if (s->dictHandle >= 0)
            s->fnSetDict(s->handle, s->dictHandle);
    }
    if (s->dictHandle < 0) return -1;

    if (mainPath)
        s->fnLoadDict(s->handle, s->dictHandle, 0, mainPath);
    if (rootPath)
        s->fnLoadDict(s->handle, s->dictHandle, 1, rootPath);

    return 0;
}
