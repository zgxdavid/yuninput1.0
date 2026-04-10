#include "TextService.h"

#include "Globals.h"

#include <Windows.h>
#include <shellapi.h>

#include <algorithm>
#include <condition_variable>
#include <cwctype>
#include <filesystem>
#include <fstream>
#include <mutex>
#include <new>
#include <regex>
#include <sstream>
#include <thread>

namespace {

constexpr const wchar_t* kBuildMarker = L"cw-r4-20260403-yuninput1-userdict-slim-v1";
constexpr ULONGLONG kAnchorReuseWindowMs = 2500ULL;
constexpr ULONGLONG kAnchorFastReuseWindowMs = 120ULL;
constexpr ULONGLONG kAutoPhraseFlushIntervalMs = 3000ULL;
constexpr ULONGLONG kUserFrequencyFlushIntervalMs = 3000ULL;
constexpr ULONGLONG kUserDataReloadMinIntervalMs = 1500ULL;
constexpr ULONGLONG kHelperAutoPhraseReloadMinIntervalMs = 500ULL;
constexpr ULONGLONG kAutoPhraseBuilderLaunchMinIntervalMs = 2000ULL;
constexpr const wchar_t* kAutoPhraseHelperWakeEventName = L"Local\\Yuninput.AutoPhraseHelper.Wake";
constexpr size_t kSessionAutoPhraseHistoryCharLimit = 2000;
constexpr size_t kSessionAutoPhraseMaxLength = 12;
constexpr size_t kRoamingAutoPhraseMaxEntries = 4096;
constexpr size_t kPrimaryCandidateQueryLimit = 96;
constexpr size_t kPrefixCandidateQueryLimit = 16;
constexpr size_t kPhraseCandidateQueryLimit = 24;
constexpr size_t kRefreshTargetCandidateCount = 18;
constexpr size_t kFastQueryScanBudgetOneChar = 384;
constexpr size_t kFastQueryScanBudgetTwoChar = 256;
constexpr wchar_t kSessionAutoPhraseBreak = static_cast<wchar_t>(0xE000);
constexpr size_t kRuntimeLogMaxQueuedLines = 512;

std::mutex g_runtimeLogMutex;
std::condition_variable g_runtimeLogCv;
std::deque<std::wstring> g_runtimeLogQueue;
std::thread g_runtimeLogThread;
bool g_runtimeLogStopRequested = false;
unsigned int g_runtimeLogClientCount = 0;

std::wstring BuildUserDataFilePath(const std::wstring& userDataDir, const wchar_t* fileName) {
    if (userDataDir.empty() || fileName == nullptr || *fileName == L'\0') {
        return L"";
    }

    return userDataDir + L"\\" + fileName;
}

bool CopyFileIfMissing(const std::filesystem::path& sourcePath, const std::filesystem::path& targetPath) {
    if (sourcePath.empty() || targetPath.empty()) {
        return false;
    }

    std::error_code ec;
    if (!std::filesystem::exists(sourcePath, ec) || std::filesystem::exists(targetPath, ec)) {
        return false;
    }

    if (targetPath.has_parent_path()) {
        std::filesystem::create_directories(targetPath.parent_path(), ec);
        ec.clear();
    }

    return std::filesystem::copy_file(sourcePath, targetPath, std::filesystem::copy_options::skip_existing, ec);
}

bool FileContainsAutoPhraseTag(const std::wstring& filePath) {
    if (filePath.empty()) {
        return false;
    }

    std::ifstream input(filePath, std::ios::in | std::ios::binary);
    if (!input) {
        return false;
    }

    std::string line;
    while (std::getline(input, line)) {
        if (line.empty() || line[0] == '#') {
            continue;
        }

        if (line.find(" auto") != std::string::npos || line.find("\tauto") != std::string::npos) {
            return true;
        }
    }

    return false;
}

bool FileContainsLegacyFrequencyNoise(const std::wstring& filePath) {
    if (filePath.empty()) {
        return false;
    }

    const auto utf8ToWide = [](const std::string& input) {
        if (input.empty()) {
            return std::wstring();
        }

        const auto decode = [&input](UINT codePage, DWORD flags) {
            const int required = MultiByteToWideChar(
                codePage,
                flags,
                input.c_str(),
                static_cast<int>(input.size()),
                nullptr,
                0);
            if (required <= 0) {
                return std::wstring();
            }

            std::wstring output(static_cast<size_t>(required), L'\0');
            const int converted = MultiByteToWideChar(
                codePage,
                flags,
                input.c_str(),
                static_cast<int>(input.size()),
                output.data(),
                required);
            if (converted <= 0) {
                return std::wstring();
            }

            return output;
        };

        std::wstring decoded = decode(CP_UTF8, MB_ERR_INVALID_CHARS);
        if (!decoded.empty()) {
            return decoded;
        }

        decoded = decode(54936, 0);
        if (!decoded.empty()) {
            return decoded;
        }

        return decode(CP_ACP, 0);
    };

    const auto isTextInGb2312 = [](const std::wstring& text) {
        if (text.empty()) {
            return true;
        }

        for (wchar_t ch : text) {
            if (ch <= 0x7F) {
                continue;
            }

            char buffer[4] = {};
            BOOL usedDefaultChar = FALSE;
            const int converted = WideCharToMultiByte(
                936,
                WC_NO_BEST_FIT_CHARS,
                &ch,
                1,
                buffer,
                static_cast<int>(sizeof(buffer)),
                nullptr,
                &usedDefaultChar);
            if (converted <= 0 || usedDefaultChar) {
                return false;
            }

            if (converted == 1) {
                if (static_cast<unsigned char>(buffer[0]) > 0x7F) {
                    return false;
                }
                continue;
            }

            if (converted != 2) {
                return false;
            }

            const unsigned char lead = static_cast<unsigned char>(buffer[0]);
            const unsigned char trail = static_cast<unsigned char>(buffer[1]);
            if (!(lead >= 0xA1 && lead <= 0xF7 && trail >= 0xA1 && trail <= 0xFE)) {
                return false;
            }
        }

        return true;
    };

    const auto isHanText = [](const std::wstring& text) {
        return !text.empty() && std::all_of(text.begin(), text.end(), [](wchar_t ch) {
            return (ch >= 0x3400 && ch <= 0x4DBF) ||
                   (ch >= 0x4E00 && ch <= 0x9FFF) ||
                   (ch >= 0xF900 && ch <= 0xFAFF);
        });
    };

    std::ifstream input(filePath, std::ios::in | std::ios::binary);
    if (!input) {
        return false;
    }

    std::string line;
    while (std::getline(input, line)) {
        if (line.empty() || line[0] == '#') {
            continue;
        }

        std::istringstream iss(line);
        std::string codeUtf8;
        std::string textUtf8;
        std::uint64_t score = 0;
        if (!(iss >> codeUtf8 >> textUtf8 >> score)) {
            return true;
        }

        const std::wstring code = utf8ToWide(codeUtf8);
        const std::wstring text = utf8ToWide(textUtf8);
        if (text.empty()) {
            return true;
        }

        const bool eligibleSingle = text.size() == 1 && isTextInGb2312(text);
        const bool eligiblePhrase = text.size() >= 2 && code.size() == 4 && isHanText(text);
        if (!eligibleSingle && !eligiblePhrase) {
            return true;
        }
    }

    return false;
}

bool IsAsciiCodeToken(const std::string& token) {
    if (token.empty()) {
        return false;
    }

    for (unsigned char ch : token) {
        if (!((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z'))) {
            return false;
        }
    }

    return true;
}

std::wstring QueryFileStampToken(const std::wstring& filePath) {
    if (filePath.empty()) {
        return L"-";
    }

    std::error_code ec;
    const std::filesystem::path path(filePath);
    const bool exists = std::filesystem::exists(path, ec);
    if (ec || !exists) {
        return L"0";
    }

    ec.clear();
    const auto writeTime = std::filesystem::last_write_time(path, ec);
    if (ec) {
        return L"1";
    }

    return L"1:" + std::to_wstring(static_cast<long long>(writeTime.time_since_epoch().count()));
}

std::wstring QuoteCommandLineArg(const std::wstring& value) {
    std::wstring quoted = L"\"";
    for (wchar_t ch : value) {
        if (ch == L'"') {
            quoted += L"\\\"";
        } else {
            quoted.push_back(ch);
        }
    }
    quoted += L'"';
    return quoted;
}

std::wstring BuildUserDataFilesStamp(
    const std::wstring& userDictPath,
    const std::wstring& autoPhraseDictPath,
    const std::wstring& userFreqPath,
    const std::wstring& blockedEntriesPath,
    const std::wstring& contextAssocPath,
    const std::wstring& contextAssocBlacklistPath) {
    return QueryFileStampToken(userDictPath) + L"|" +
        QueryFileStampToken(autoPhraseDictPath) + L"|" +
        QueryFileStampToken(userFreqPath) + L"|" +
        QueryFileStampToken(blockedEntriesPath) + L"|" +
        QueryFileStampToken(contextAssocPath) + L"|" +
        QueryFileStampToken(contextAssocBlacklistPath);
}

bool SignalAutoPhraseHelperWakeEvent() {
    HANDLE wakeEvent = OpenEventW(EVENT_MODIFY_STATE, FALSE, kAutoPhraseHelperWakeEventName);
    if (wakeEvent == nullptr) {
        return false;
    }

    const BOOL signaled = SetEvent(wakeEvent);
    CloseHandle(wakeEvent);
    return signaled == TRUE;
}

void MigrateLegacyUserDataFiles(const std::wstring& userDataDir) {
    wchar_t localAppData[MAX_PATH] = {};
    const DWORD localLen = GetEnvironmentVariableW(L"LOCALAPPDATA", localAppData, MAX_PATH);
    if (localLen == 0 || localLen >= MAX_PATH) {
        return;
    }

    std::error_code ec;
    const std::filesystem::path targetDir(userDataDir);
    const std::filesystem::path legacyDir = std::filesystem::path(std::wstring(localAppData, localLen)) / L"yuninput";
    if (!std::filesystem::exists(legacyDir, ec)) {
        return;
    }

    ec.clear();
    if (std::filesystem::equivalent(legacyDir, targetDir, ec)) {
        return;
    }

    CopyFileIfMissing(legacyDir / L"yuninput_user.dict", targetDir / L"yuninput_user.dict");
    CopyFileIfMissing(legacyDir / L"user_dict.txt", targetDir / L"yuninput_user.dict");
    CopyFileIfMissing(legacyDir / L"user_freq.txt", targetDir / L"user_freq.txt");
    CopyFileIfMissing(legacyDir / L"blocked_entries.txt", targetDir / L"blocked_entries.txt");
    CopyFileIfMissing(legacyDir / L"context_assoc.txt", targetDir / L"context_assoc.txt");
    CopyFileIfMissing(legacyDir / L"context_assoc_blacklist.txt", targetDir / L"context_assoc_blacklist.txt");
    CopyFileIfMissing(legacyDir / L"manual_phrase_review.txt", targetDir / L"manual_phrase_review.txt");
    CopyFileIfMissing(legacyDir / L"auto_phrase_session.txt", targetDir / L"auto_phrase_session.txt");
}

std::string WideToUtf8ForLog(const std::wstring& input) {
    if (input.empty()) {
        return "";
    }

    const int required = WideCharToMultiByte(CP_UTF8, 0, input.c_str(), static_cast<int>(input.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0) {
        return "";
    }

    std::string output(static_cast<size_t>(required), '\0');
    WideCharToMultiByte(CP_UTF8, 0, input.c_str(), static_cast<int>(input.size()), output.data(), required, nullptr, nullptr);
    return output;
}

std::string WideToUtf8Text(const std::wstring& input) {
    if (input.empty()) {
        return "";
    }

    const int required = WideCharToMultiByte(CP_UTF8, 0, input.c_str(), static_cast<int>(input.size()), nullptr, 0, nullptr, nullptr);
    if (required <= 0) {
        return "";
    }

    std::string output(static_cast<size_t>(required), '\0');
    WideCharToMultiByte(CP_UTF8, 0, input.c_str(), static_cast<int>(input.size()), output.data(), required, nullptr, nullptr);
    return output;
}

std::wstring Utf8ToWideText(const std::string& input) {
    if (input.empty()) {
        return L"";
    }

    const auto decode = [&input](UINT codePage, DWORD flags) {
        const int required = MultiByteToWideChar(
            codePage,
            flags,
            input.c_str(),
            static_cast<int>(input.size()),
            nullptr,
            0);
        if (required <= 0) {
            return std::wstring();
        }

        std::wstring output(static_cast<size_t>(required), L'\0');
        const int converted = MultiByteToWideChar(
            codePage,
            flags,
            input.c_str(),
            static_cast<int>(input.size()),
            output.data(),
            required);
        if (converted <= 0) {
            return std::wstring();
        }
        return output;
    };

    std::wstring decoded = decode(CP_UTF8, MB_ERR_INVALID_CHARS);
    if (!decoded.empty()) {
        return decoded;
    }

    decoded = decode(54936, 0);  // GB18030 fallback for legacy dictionary files.
    if (!decoded.empty()) {
        return decoded;
    }

    return decode(CP_ACP, 0);
}

bool IsCharInGB2312(wchar_t ch) {
    if (ch <= 0x7F) {
        return true;
    }

    char buffer[4] = {};
    BOOL usedDefaultChar = FALSE;
    const int converted = WideCharToMultiByte(
        936,
        WC_NO_BEST_FIT_CHARS,
        &ch,
        1,
        buffer,
        static_cast<int>(sizeof(buffer)),
        nullptr,
        &usedDefaultChar);
    if (converted <= 0 || usedDefaultChar) {
        return false;
    }

    if (converted == 1) {
        return static_cast<unsigned char>(buffer[0]) <= 0x7F;
    }

    if (converted != 2) {
        return false;
    }

    const unsigned char lead = static_cast<unsigned char>(buffer[0]);
    const unsigned char trail = static_cast<unsigned char>(buffer[1]);
    return lead >= 0xA1 && lead <= 0xF7 && trail >= 0xA1 && trail <= 0xFE;
}

bool IsTextInGB2312(const std::wstring& text) {
    if (text.empty()) {
        return true;
    }

    for (wchar_t ch : text) {
        if (!IsCharInGB2312(ch)) {
            return false;
        }
    }
    return true;
}

std::wstring GetRuntimeLogPath() {
    wchar_t localAppData[MAX_PATH] = {};
    const DWORD len = GetEnvironmentVariableW(L"LOCALAPPDATA", localAppData, MAX_PATH);
    if (len == 0 || len >= MAX_PATH) {
        return L"C:\\Windows\\Temp\\yuninput_runtime.log";
    }

    return std::wstring(localAppData, len) + L"\\yuninput\\runtime.log";
}

bool IsTruthyEnvironmentValue(const wchar_t* value) {
    if (value == nullptr || *value == L'\0') {
        return false;
    }

    return _wcsicmp(value, L"1") == 0 ||
           _wcsicmp(value, L"true") == 0 ||
           _wcsicmp(value, L"yes") == 0 ||
           _wcsicmp(value, L"on") == 0;
}

bool IsVerboseKeyTraceEnabled() {
#ifdef _DEBUG
    return true;
#else
    static const bool enabled = []() {
        if (IsDebuggerPresent()) {
            return true;
        }

        wchar_t buffer[16] = {};
        const DWORD len = GetEnvironmentVariableW(L"YUNINPUT_VERBOSE_KEY_TRACE", buffer, _countof(buffer));
        if (len == 0 || len >= _countof(buffer)) {
            return false;
        }

        return IsTruthyEnvironmentValue(buffer);
    }();
    return enabled;
#endif
}

void Trace(const std::wstring& text);

bool IsLatencyProfilingEnabled() {
#ifdef _DEBUG
    return true;
#else
    static const bool enabled = []() {
        wchar_t buffer[16] = {};
        const DWORD len = GetEnvironmentVariableW(L"YUNINPUT_DISABLE_LATENCY_TRACE", buffer, _countof(buffer));
        if (len == 0 || len >= _countof(buffer)) {
            return true;
        }

        return !IsTruthyEnvironmentValue(buffer);
    }();
    return enabled;
#endif
}

LONGLONG QueryPerfCounterValue() {
    LARGE_INTEGER value = {};
    QueryPerformanceCounter(&value);
    return value.QuadPart;
}

double PerfCounterElapsedMs(LONGLONG startCounter, LONGLONG endCounter) {
    static const double frequency = []() {
        LARGE_INTEGER value = {};
        QueryPerformanceFrequency(&value);
        return value.QuadPart > 0 ? static_cast<double>(value.QuadPart) : 0.0;
    }();

    if (frequency <= 0.0) {
        return 0.0;
    }

    return (static_cast<double>(endCounter - startCounter) * 1000.0) / frequency;
}

void TraceLatencySample(const std::wstring& label, LONGLONG startCounter, LONGLONG endCounter) {
    if (!IsLatencyProfilingEnabled()) {
        return;
    }

    wchar_t elapsedBuffer[32] = {};
    swprintf_s(elapsedBuffer, L"%.3f", PerfCounterElapsedMs(startCounter, endCounter));
    Trace(L"latency " + label + L" ms=" + elapsedBuffer);
}

void WriteRuntimeLogBatch(const std::deque<std::wstring>& batch) {
    if (batch.empty()) {
        return;
    }

    const std::wstring logPath = GetRuntimeLogPath();
    const std::size_t slash = logPath.find_last_of(L"\\/");
    if (slash != std::wstring::npos) {
        const std::wstring dirPath = logPath.substr(0, slash);
        CreateDirectoryW(dirPath.c_str(), nullptr);
    }

    std::ofstream stream(logPath, std::ios::app | std::ios::binary);
    if (!stream.is_open()) {
        return;
    }

    SYSTEMTIME now = {};
    wchar_t prefix[64] = {};
    for (const std::wstring& message : batch) {
        GetLocalTime(&now);
        swprintf_s(
            prefix,
            L"[%04u-%02u-%02u %02u:%02u:%02u] ",
            now.wYear,
            now.wMonth,
            now.wDay,
            now.wHour,
            now.wMinute,
            now.wSecond);

        const std::wstring line = std::wstring(prefix) + message + L"\r\n";
        const std::string utf8 = WideToUtf8ForLog(line);
        stream.write(utf8.data(), static_cast<std::streamsize>(utf8.size()));
    }
}

void RuntimeLogWorkerMain() {
    for (;;) {
        std::deque<std::wstring> batch;
        {
            std::unique_lock<std::mutex> lock(g_runtimeLogMutex);
            g_runtimeLogCv.wait(lock, []() {
                return g_runtimeLogStopRequested || !g_runtimeLogQueue.empty();
            });

            if (g_runtimeLogStopRequested && g_runtimeLogQueue.empty()) {
                break;
            }

            batch.swap(g_runtimeLogQueue);
        }

        WriteRuntimeLogBatch(batch);
    }
}

void StartRuntimeLogWorker() {
    std::lock_guard<std::mutex> lock(g_runtimeLogMutex);
    ++g_runtimeLogClientCount;
    if (g_runtimeLogThread.joinable()) {
        return;
    }

    g_runtimeLogStopRequested = false;
    g_runtimeLogThread = std::thread(&RuntimeLogWorkerMain);
}

void StopRuntimeLogWorker() {
    std::thread worker;
    {
        std::lock_guard<std::mutex> lock(g_runtimeLogMutex);
        if (g_runtimeLogClientCount > 0) {
            --g_runtimeLogClientCount;
        }

        if (g_runtimeLogClientCount != 0 || !g_runtimeLogThread.joinable()) {
            return;
        }

        g_runtimeLogStopRequested = true;
        worker = std::move(g_runtimeLogThread);
    }

    g_runtimeLogCv.notify_all();
    worker.join();
}

void AppendRuntimeLog(const std::wstring& message) {
    std::lock_guard<std::mutex> lock(g_runtimeLogMutex);
    if (!g_runtimeLogThread.joinable()) {
        std::deque<std::wstring> singleEntry;
        singleEntry.push_back(message);
        WriteRuntimeLogBatch(singleEntry);
        return;
    }

    if (g_runtimeLogQueue.size() >= kRuntimeLogMaxQueuedLines) {
        g_runtimeLogQueue.pop_front();
    }
    g_runtimeLogQueue.push_back(message);
    g_runtimeLogCv.notify_all();
}

void Trace(const std::wstring& text) {
    OutputDebugStringW((L"[yuninput] " + text + L"\n").c_str());
    AppendRuntimeLog(text);
}

WPARAM NormalizeVirtualKey(WPARAM wParam, LPARAM lParam) {
    if (wParam != VK_PROCESSKEY) {
        return wParam;
    }

    const UINT scanCode = (static_cast<UINT>(lParam) >> 16) & 0xFF;
    const bool extended = ((static_cast<UINT>(lParam) >> 24) & 0x1) != 0;
    const UINT vsc = extended ? (scanCode | 0xE000) : scanCode;
    const UINT mapped = MapVirtualKeyW(vsc, MAPVK_VSC_TO_VK_EX);
    if (mapped != 0) {
        return static_cast<WPARAM>(mapped);
    }

    return wParam;
}

bool IsLeftShiftToggleKey(WPARAM key, LPARAM lParam) {
    if (key == VK_LSHIFT) {
        return true;
    }

    if (key != VK_SHIFT) {
        return false;
    }

    const UINT scanCode = (static_cast<UINT>(lParam) >> 16) & 0xFF;
    const UINT leftShiftScanCode = MapVirtualKeyW(VK_LSHIFT, MAPVK_VK_TO_VSC);
    if (scanCode != 0) {
        return scanCode == leftShiftScanCode || scanCode == 0x2A;
    }

    const bool leftDown =
        (GetKeyState(VK_LSHIFT) & 0x8000) != 0 ||
        (GetAsyncKeyState(VK_LSHIFT) & 0x8000) != 0;
    const bool rightDown =
        (GetKeyState(VK_RSHIFT) & 0x8000) != 0 ||
        (GetAsyncKeyState(VK_RSHIFT) & 0x8000) != 0;

    if (leftDown && !rightDown) {
        return true;
    }

    // Some TSF hosts provide VK_SHIFT without reliable scan codes/key-state split.
    // Prefer enabling the toggle path instead of silently missing the key.
    return true;
}

bool TryGetCaretScreenPoint(POINT& outPoint) {
    const int virtualLeft = GetSystemMetrics(SM_XVIRTUALSCREEN);
    const int virtualTop = GetSystemMetrics(SM_YVIRTUALSCREEN);
    const int virtualWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
    const int virtualHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN);
    const int virtualRight = virtualLeft + std::max(virtualWidth, 1);
    const int virtualBottom = virtualTop + std::max(virtualHeight, 1);

    const auto isPointReasonable = [&](const POINT& pt) {
        return pt.x >= virtualLeft - 128 &&
               pt.x <= virtualRight + 128 &&
               pt.y >= virtualTop - 128 &&
               pt.y <= virtualBottom + 128;
    };

    GUITHREADINFO info = {};
    info.cbSize = sizeof(info);
    DWORD guiThreadId = 0;
    HWND foreground = GetForegroundWindow();
    if (foreground != nullptr) {
        guiThreadId = GetWindowThreadProcessId(foreground, nullptr);
    }

    if (GetGUIThreadInfo(guiThreadId, &info)) {
        if (!IsRectEmpty(&info.rcCaret)) {
            POINT pt = {info.rcCaret.left, info.rcCaret.bottom};
            HWND anchorHwnd = info.hwndCaret != nullptr ? info.hwndCaret : info.hwndFocus;
            if (anchorHwnd != nullptr && ClientToScreen(anchorHwnd, &pt) && isPointReasonable(pt)) {
                outPoint = pt;
                return true;
            }

            if (anchorHwnd == nullptr && isPointReasonable(pt)) {
                outPoint = pt;
                return true;
            }
        }

        if (info.hwndFocus != nullptr) {
            RECT focusRect = {};
            if (GetWindowRect(info.hwndFocus, &focusRect) && !IsRectEmpty(&focusRect)) {
                const int focusWidth = static_cast<int>(focusRect.right - focusRect.left);
                const int focusHeight = static_cast<int>(focusRect.bottom - focusRect.top);
                const int anchorOffsetX = std::min(48, std::max(16, focusWidth / 6));
                const int anchorOffsetY = std::min(160, std::max(28, focusHeight / 3));
                POINT pt = {
                    static_cast<LONG>(focusRect.left + anchorOffsetX),
                    static_cast<LONG>(focusRect.top + anchorOffsetY)};
                if (isPointReasonable(pt)) {
                    outPoint = pt;
                    return true;
                }
            }
        }
    }

    POINT cursor = {};
    if (GetCursorPos(&cursor)) {
        // Cursor-based fallback keeps the candidate near current typing interaction.
        outPoint = cursor;
        return true;
    }

    POINT pt = {};
    HWND focus = GetFocus();
    if (focus != nullptr && GetCaretPos(&pt) && ClientToScreen(focus, &pt)) {
        outPoint = pt;
        return true;
    }

    return false;
}

bool TryGetFocusedContext(ITfThreadMgr* threadMgr, ITfContext** outContext) {
    if (outContext == nullptr) {
        return false;
    }

    *outContext = nullptr;
    if (threadMgr == nullptr) {
        return false;
    }

    ITfDocumentMgr* documentMgr = nullptr;
    const HRESULT focusHr = threadMgr->GetFocus(&documentMgr);
    if (FAILED(focusHr) || documentMgr == nullptr) {
        return false;
    }

    const HRESULT topHr = documentMgr->GetTop(outContext);
    documentMgr->Release();
    return SUCCEEDED(topHr) && *outContext != nullptr;
}

bool TryMapSelectionKeyToIndex(WPARAM key, size_t& outIndex) {
    if (key >= L'1' && key <= L'9') {
        outIndex = static_cast<size_t>(key - L'1');
        return true;
    }

    if (key >= VK_NUMPAD1 && key <= VK_NUMPAD9) {
        outIndex = static_cast<size_t>(key - VK_NUMPAD1);
        return true;
    }

    return false;
}

bool TryGetTypedChar(WPARAM wParam, LPARAM lParam, wchar_t& outChar) {
    BYTE keyboardState[256] = {};
    if (!GetKeyboardState(keyboardState)) {
        return false;
    }

    WCHAR chars[4] = {};
    const UINT scanCode = (static_cast<UINT>(lParam) >> 16) & 0xFF;
    const HKL layout = GetKeyboardLayout(0);
    const int count = ToUnicodeEx(static_cast<UINT>(wParam), scanCode, keyboardState, chars, 4, 0, layout);
    if (count == 1) {
        outChar = chars[0];
        return true;
    }

    if (count < 0) {
        WCHAR clearChars[4] = {};
        ToUnicodeEx(static_cast<UINT>(wParam), scanCode, keyboardState, clearChars, 4, 0, layout);
    }

    return false;
}

bool IsAsciiAlphaVirtualKey(WPARAM wParam) {
    return (wParam >= L'A' && wParam <= L'Z') || (wParam >= L'a' && wParam <= L'z');
}

bool TryBuildDirectAsciiCommitChar(WPARAM wParam, LPARAM lParam, wchar_t& outChar) {
    wchar_t translated = 0;
    if (TryGetTypedChar(wParam, lParam, translated)) {
        if (translated >= 0x20 && translated != 0x7F) {
            outChar = translated;
            return true;
        }
        return false;
    }

    wchar_t fallback = 0;
    if (wParam == VK_SPACE) {
        fallback = L' ';
    } else if (IsAsciiAlphaVirtualKey(wParam)) {
        fallback = static_cast<wchar_t>(wParam);
        const bool shiftPressed = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        const bool capsOn = (GetKeyState(VK_CAPITAL) & 0x0001) != 0;
        const bool useUpper = shiftPressed != capsOn;
        fallback = useUpper
            ? static_cast<wchar_t>(std::towupper(static_cast<wint_t>(fallback)))
            : static_cast<wchar_t>(std::towlower(static_cast<wint_t>(fallback)));
    } else if (wParam >= L'0' && wParam <= L'9') {
        fallback = static_cast<wchar_t>(wParam);
    } else if (wParam >= VK_NUMPAD0 && wParam <= VK_NUMPAD9) {
        fallback = static_cast<wchar_t>(L'0' + (wParam - VK_NUMPAD0));
    } else if (wParam == VK_OEM_MINUS) {
        const bool shiftPressed = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        fallback = shiftPressed ? L'_' : L'-';
    } else if (wParam == VK_OEM_PLUS) {
        const bool shiftPressed = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
        fallback = shiftPressed ? L'+' : L'=';
    } else {
        return false;
    }

    outChar = fallback;
    return true;
}

bool TryMapSelectionInputToIndex(WPARAM key, LPARAM lParam, size_t& outIndex) {
    if (!TryMapSelectionKeyToIndex(key, outIndex)) {
        return false;
    }

    if (key >= L'1' && key <= L'9') {
        wchar_t input = 0;
        if (TryGetTypedChar(key, lParam, input) && std::iswdigit(static_cast<wint_t>(input)) == 0) {
            return false;
        }
    }

    return true;
}

std::wstring TranslateChinesePunctuation(
    wchar_t input,
    bool smartSymbolPairs,
    bool& nextSingleQuoteOpen,
    bool& nextDoubleQuoteOpen) {
    switch (input) {
    case L',':
        return L"\uFF0C";
    case L'.':
        return L"\u3002";
    case L'/':
        return L"\u3001";
    case L'?':
        return L"\uFF1F";
    case L';':
        return L"\uFF1B";
    case L':':
        return L"\uFF1A";
    case L'!':
        return L"\uFF01";
    case L'(':
        return L"\uFF08";
    case L')':
        return L"\uFF09";
    case L'[':
        return smartSymbolPairs ? L"\u3010" : L"\uFF3B";
    case L']':
        return smartSymbolPairs ? L"\u3011" : L"\uFF3D";
    case L'{':
        return L"\uFF5B";
    case L'}':
        return L"\uFF5D";
    case L'<':
        return smartSymbolPairs ? L"\u300A" : L"\u3008";
    case L'>':
        return smartSymbolPairs ? L"\u300B" : L"\u3009";
    case L'\\':
        return L"\u3001";
    case L'-':
        return L"\uFF0D";
    case L'_':
        return L"\u2014\u2014";
    case L'~':
        return L"\uFF5E";
    case L'=':
        return L"\uFF1D";
    case L'+':
        return L"\uFF0B";
    case L'\'': {
        if (!smartSymbolPairs) {
            return L"\uFF07";
        }
        const std::wstring output = nextSingleQuoteOpen ? L"\u2018" : L"\u2019";
        nextSingleQuoteOpen = !nextSingleQuoteOpen;
        return output;
    }
    case L'\"': {
        if (!smartSymbolPairs) {
            return L"\uFF02";
        }
        const std::wstring output = nextDoubleQuoteOpen ? L"\u201C" : L"\u201D";
        nextDoubleQuoteOpen = !nextDoubleQuoteOpen;
        return output;
    }
    default:
        return L"";
    }
}

bool TryTranslateChinesePunctuation(
    WPARAM wParam,
    LPARAM lParam,
    bool chinesePunctuation,
    bool smartSymbolPairs,
    bool& nextSingleQuoteOpen,
    bool& nextDoubleQuoteOpen,
    std::wstring& outText) {
    outText.clear();
    if (!chinesePunctuation) {
        return false;
    }

    wchar_t input = 0;
    if (!TryGetTypedChar(wParam, lParam, input)) {
        return false;
    }

    outText = TranslateChinesePunctuation(input, smartSymbolPairs, nextSingleQuoteOpen, nextDoubleQuoteOpen);
    return !outText.empty();
}

bool IsPunctuationVirtualKey(WPARAM key) {
    return key == VK_OEM_1 ||
           key == VK_OEM_2 ||
           key == VK_OEM_3 ||
           key == VK_OEM_4 ||
           key == VK_OEM_5 ||
           key == VK_OEM_6 ||
           key == VK_OEM_7 ||
           key == VK_OEM_COMMA ||
           key == VK_OEM_MINUS ||
           key == VK_SUBTRACT ||
           key == VK_OEM_PERIOD ||
           key == VK_OEM_PLUS ||
           key == VK_ADD;
}

bool IsMinusPageUpKey(WPARAM key, WPARAM rawKey, LPARAM lParam) {
    if (key == VK_OEM_MINUS || key == VK_SUBTRACT || rawKey == VK_OEM_MINUS || rawKey == VK_SUBTRACT) {
        return true;
    }

    wchar_t input = 0;
    if (!TryGetTypedChar(rawKey, lParam, input)) {
        return false;
    }

    return input == L'-' || input == L'_';
}

bool IsPlusPageDownKey(WPARAM key, WPARAM rawKey, LPARAM lParam) {
    if (key == VK_OEM_PLUS || key == VK_ADD || rawKey == VK_OEM_PLUS || rawKey == VK_ADD) {
        return true;
    }

    wchar_t input = 0;
    if (!TryGetTypedChar(rawKey, lParam, input)) {
        return false;
    }

    return input == L'=' || input == L'+';
}

bool TryBuildPunctuationCommitText(
    WPARAM wParam,
    LPARAM lParam,
    bool chinesePunctuation,
    bool smartSymbolPairs,
    bool& nextSingleQuoteOpen,
    bool& nextDoubleQuoteOpen,
    std::wstring& outText) {
    outText.clear();

    if (TryTranslateChinesePunctuation(
            wParam,
            lParam,
            chinesePunctuation,
            smartSymbolPairs,
            nextSingleQuoteOpen,
            nextDoubleQuoteOpen,
            outText)) {
        return true;
    }

    wchar_t input = 0;
    if (!TryGetTypedChar(wParam, lParam, input)) {
        return false;
    }

    if (std::iswalpha(static_cast<wint_t>(input)) != 0 ||
        std::iswdigit(static_cast<wint_t>(input)) != 0 ||
        std::iswspace(static_cast<wint_t>(input)) != 0) {
        return false;
    }

    outText.assign(1, input);
    return true;
}

constexpr UINT kMenuToggleChinese = 1001;
constexpr UINT kMenuToggleShape = 1002;
constexpr UINT kMenuHelp = 1003;
constexpr UINT kMenuOpenConfig = 1004;
constexpr UINT kMenuOpenSystemSettings = 1005;
constexpr UINT kMenuOpenRuntimeLog = 1006;

std::filesystem::path ResolveDataRootFromModule(const std::wstring& modulePath);

const GUID GUID_LBI_YUNINPUT_CONFIG = {
    0x5988f595,
    0xf19f,
    0x4ced,
    {0x8a, 0x4e, 0xfb, 0x58, 0x97, 0x5b, 0x72, 0x43}};

std::wstring GetLocalAppDataPath() {
    wchar_t localAppData[MAX_PATH] = {};
    const DWORD localLen = GetEnvironmentVariableW(L"LOCALAPPDATA", localAppData, MAX_PATH);
    if (localLen > 0 && localLen < MAX_PATH) {
        return std::wstring(localAppData, localLen);
    }

    return L"";
}

std::wstring GetInstalledConfigExePath() {
    const std::wstring localAppData = GetLocalAppDataPath();
    if (localAppData.empty()) {
        return L"";
    }

    return localAppData + L"\\yuninput\\yuninput_config.exe";
}

bool LaunchConfigExecutable() {
    std::wstring modulePath;
    if (!GetModulePath(modulePath)) {
        return false;
    }

    const auto root = ResolveDataRootFromModule(modulePath);
    const auto fallbackPath = (root / L"yuninput_config.exe").wstring();
    if (std::filesystem::exists(fallbackPath)) {
        const HINSTANCE hInst = ShellExecuteW(nullptr, L"open", fallbackPath.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
        return reinterpret_cast<INT_PTR>(hInst) > 32;
    }

    const std::wstring installedPath = GetInstalledConfigExePath();
    if (!installedPath.empty() && std::filesystem::exists(installedPath)) {
        const HINSTANCE hInst = ShellExecuteW(nullptr, L"open", installedPath.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
        return reinterpret_cast<INT_PTR>(hInst) > 32;
    }

    return false;
}

bool LaunchSystemImeSettings() {
    const wchar_t* uris[] = {
        L"ms-settings:typing",
        L"ms-settings:regionlanguage",
    };

    for (const auto* uri : uris) {
        const HINSTANCE hInst = ShellExecuteW(nullptr, L"open", uri, nullptr, nullptr, SW_SHOWNORMAL);
        if (reinterpret_cast<INT_PTR>(hInst) > 32) {
            return true;
        }
    }

    const HINSTANCE controlInst = ShellExecuteW(nullptr, L"open", L"control.exe", L"/name Microsoft.Language", nullptr, SW_SHOWNORMAL);
    return reinterpret_cast<INT_PTR>(controlInst) > 32;
}

bool LaunchRuntimeLogFile() {
    const std::wstring logPath = GetRuntimeLogPath();
    const std::size_t slash = logPath.find_last_of(L"\\/");
    if (slash != std::wstring::npos) {
        const std::wstring dirPath = logPath.substr(0, slash);
        CreateDirectoryW(dirPath.c_str(), nullptr);
    }

    std::ofstream touch(logPath, std::ios::app | std::ios::binary);
    if (!touch.is_open()) {
        return false;
    }
    touch.close();

    const HINSTANCE hInst = ShellExecuteW(nullptr, L"open", logPath.c_str(), nullptr, nullptr, SW_SHOWNORMAL);
    return reinterpret_cast<INT_PTR>(hInst) > 32;
}

class ConfigLangBarButton final : public ITfLangBarItemButton {
public:
    ConfigLangBarButton() : refCount_(1) {
        info_.clsidService = CLSID_YunmaTextService;
        info_.guidItem = GUID_LBI_YUNINPUT_CONFIG;
        info_.dwStyle = TF_LBI_STYLE_SHOWNINTRAY | TF_LBI_STYLE_TEXTCOLORICON;
        info_.ulSort = 0;
        wcscpy_s(info_.szDescription, L"Yuninput Quick Config");
    }

    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) override {
        if (ppvObj == nullptr) {
            return E_INVALIDARG;
        }

        *ppvObj = nullptr;
        if (riid == IID_IUnknown || riid == IID_ITfLangBarItem || riid == IID_ITfLangBarItemButton) {
            *ppvObj = static_cast<ITfLangBarItemButton*>(this);
            AddRef();
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    STDMETHODIMP_(ULONG) AddRef() override {
        return static_cast<ULONG>(InterlockedIncrement(&refCount_));
    }

    STDMETHODIMP_(ULONG) Release() override {
        const LONG value = InterlockedDecrement(&refCount_);
        if (value == 0) {
            delete this;
            return 0;
        }

        return static_cast<ULONG>(value);
    }

    STDMETHODIMP GetInfo(TF_LANGBARITEMINFO* info) override {
        if (info == nullptr) {
            return E_INVALIDARG;
        }

        *info = info_;
        return S_OK;
    }

    STDMETHODIMP GetStatus(DWORD* status) override {
        if (status == nullptr) {
            return E_INVALIDARG;
        }

        *status = 0;
        return S_OK;
    }

    STDMETHODIMP Show(BOOL) override {
        return S_OK;
    }

    STDMETHODIMP GetTooltipString(BSTR* tooltip) override {
        if (tooltip == nullptr) {
            return E_INVALIDARG;
        }

        *tooltip = SysAllocString(L"Left click: open full config; Right click: quick menu");
        return (*tooltip != nullptr) ? S_OK : E_OUTOFMEMORY;
    }

    STDMETHODIMP OnClick(TfLBIClick click, POINT, const RECT*) override {
        if (click == TF_LBI_CLK_LEFT) {
            const bool launched = LaunchConfigExecutable();
            Trace(std::wstring(L"langbar left(full config)=") + (launched ? L"1" : L"0"));
        }
        else if (click == TF_LBI_CLK_RIGHT) {
            POINT pt = {};
            if (!GetCursorPos(&pt)) {
                pt.x = 0;
                pt.y = 0;
            }

            HMENU menu = CreatePopupMenu();
            if (menu != nullptr) {
                AppendMenuW(menu, MF_STRING, kMenuOpenConfig, L"Open Full Config...");
                AppendMenuW(menu, MF_STRING, kMenuOpenRuntimeLog, L"Open Runtime Log...");
                AppendMenuW(menu, MF_STRING, kMenuOpenSystemSettings, L"Open System Input Settings...");
                HWND menuOwner = GetForegroundWindow();
                if (menuOwner == nullptr) {
                    menuOwner = GetDesktopWindow();
                }
                SetForegroundWindow(menuOwner);
                const UINT cmd = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_NONOTIFY | TPM_RIGHTBUTTON, pt.x, pt.y, 0, menuOwner, nullptr);
                DestroyMenu(menu);

                if (cmd == kMenuOpenConfig) {
                    const bool launched = LaunchConfigExecutable();
                    Trace(std::wstring(L"langbar quick menu open config=") + (launched ? L"1" : L"0"));
                }
                else if (cmd == kMenuOpenRuntimeLog) {
                    const bool launched = LaunchRuntimeLogFile();
                    Trace(std::wstring(L"langbar quick menu open runtime log=") + (launched ? L"1" : L"0"));
                }
                else if (cmd == kMenuOpenSystemSettings) {
                    const bool launched = LaunchSystemImeSettings();
                    Trace(std::wstring(L"langbar quick menu open system settings=") + (launched ? L"1" : L"0"));
                }
            }
        }
        return S_OK;
    }

    STDMETHODIMP InitMenu(ITfMenu* menu) override {
        (void)menu;
        return E_NOTIMPL;
    }

    STDMETHODIMP OnMenuSelect(UINT wID) override {
        if (wID == kMenuOpenConfig) {
            LaunchConfigExecutable();
        }
        else if (wID == kMenuOpenSystemSettings) {
            LaunchSystemImeSettings();
        }
        else if (wID == kMenuOpenRuntimeLog) {
            LaunchRuntimeLogFile();
        }
        return S_OK;
    }

    STDMETHODIMP GetIcon(HICON* icon) override {
        if (icon == nullptr) {
            return E_INVALIDARG;
        }

        *icon = static_cast<HICON>(LoadImageW(g_moduleHandle, MAKEINTRESOURCEW(1), IMAGE_ICON, 16, 16, LR_DEFAULTCOLOR));
        if (*icon == nullptr) {
            *icon = LoadIconW(nullptr, IDI_APPLICATION);
        }
        return (*icon != nullptr) ? S_OK : E_FAIL;
    }

    STDMETHODIMP GetText(BSTR* text) override {
        if (text == nullptr) {
            return E_INVALIDARG;
        }

        *text = SysAllocString(L"\u5300");
        return (*text != nullptr) ? S_OK : E_OUTOFMEMORY;
    }

private:
    ~ConfigLangBarButton() = default;

    LONG refCount_;
    TF_LANGBARITEMINFO info_ = {};
};

bool SendUnicodeTextWithInput(const std::wstring& text) {
    if (text.empty()) {
        return false;
    }

    std::vector<INPUT> inputs;
    inputs.reserve(text.size() * 2);

    for (wchar_t ch : text) {
        INPUT down = {};
        down.type = INPUT_KEYBOARD;
        down.ki.wScan = ch;
        down.ki.dwFlags = KEYEVENTF_UNICODE;
        inputs.push_back(down);

        INPUT up = down;
        up.ki.dwFlags = KEYEVENTF_UNICODE | KEYEVENTF_KEYUP;
        inputs.push_back(up);
    }

    const UINT sent = SendInput(static_cast<UINT>(inputs.size()), inputs.data(), sizeof(INPUT));
    return sent == inputs.size();
}

bool SafeRequestEditSession(
    ITfContext* context,
    TfClientId clientId,
    ITfEditSession* editSession,
    HRESULT& outRequestHr,
    HRESULT& outEditHr) {
    if (context == nullptr || editSession == nullptr) {
        outRequestHr = E_INVALIDARG;
        outEditHr = E_INVALIDARG;
        return false;
    }

    outRequestHr = E_FAIL;
    outEditHr = E_FAIL;

    __try {
        outRequestHr = context->RequestEditSession(
            clientId,
            editSession,
            TF_ES_ASYNCDONTCARE | TF_ES_READWRITE,
            &outEditHr);
        return true;
    }
    __except (EXCEPTION_EXECUTE_HANDLER) {
        outRequestHr = E_UNEXPECTED;
        outEditHr = E_UNEXPECTED;
        return false;
    }
}

std::filesystem::path ResolveDataRootFromModule(const std::wstring& modulePath) {
    const auto moduleFile = std::filesystem::path(modulePath);
    const auto moduleDir = moduleFile.parent_path();
    const auto dirName = moduleDir.filename().wstring();

    if (_wcsicmp(dirName.c_str(), L"bin") == 0) {
        return moduleDir.parent_path();
    }

    if (_wcsicmp(dirName.c_str(), L"Release") == 0 ||
        _wcsicmp(dirName.c_str(), L"Debug") == 0 ||
        _wcsicmp(dirName.c_str(), L"RelWithDebInfo") == 0 ||
        _wcsicmp(dirName.c_str(), L"MinSizeRel") == 0) {
        const auto buildDir = moduleDir.parent_path();
        if (_wcsicmp(buildDir.filename().wstring().c_str(), L"build") == 0) {
            return buildDir.parent_path();
        }
    }

    return moduleDir.parent_path();
}

std::wstring ReadTextFile(const std::filesystem::path& path) {
    std::ifstream input(path, std::ios::in | std::ios::binary);
    if (!input) {
        return L"";
    }

    std::string bytes((std::istreambuf_iterator<char>(input)), std::istreambuf_iterator<char>());
    if (bytes.empty()) {
        return L"";
    }

    if (bytes.size() >= 3 &&
        static_cast<unsigned char>(bytes[0]) == 0xEF &&
        static_cast<unsigned char>(bytes[1]) == 0xBB &&
        static_cast<unsigned char>(bytes[2]) == 0xBF) {
        bytes.erase(0, 3);
    }

    return Utf8ToWideText(bytes);
}

bool ExtractBool(const std::wstring& text, const std::wstring& key, bool& outValue) {
    const std::wregex re(L"\"" + key + L"\"\\s*:\\s*(true|false)", std::regex_constants::icase);
    std::wsmatch match;
    if (!std::regex_search(text, match, re) || match.size() < 2) {
        return false;
    }

    outValue = (_wcsicmp(match[1].str().c_str(), L"true") == 0);
    return true;
}

bool ExtractInt(const std::wstring& text, const std::wstring& key, int& outValue) {
    const std::wregex re(L"\"" + key + L"\"\\s*:\\s*(-?[0-9]+)");
    std::wsmatch match;
    if (!std::regex_search(text, match, re) || match.size() < 2) {
        return false;
    }

    outValue = _wtoi(match[1].str().c_str());
    return true;
}

bool ExtractString(const std::wstring& text, const std::wstring& key, std::wstring& outValue) {
    const std::wregex re(L"\"" + key + L"\"\\s*:\\s*\"([^\"]*)\"");
    std::wsmatch match;
    if (!std::regex_search(text, match, re) || match.size() < 2) {
        return false;
    }

    outValue = match[1].str();
    return true;
}

class InsertTextEditSession final : public ITfEditSession {
public:
    InsertTextEditSession(ITfContext* context, std::wstring text) : refCount_(1), context_(context), text_(std::move(text)) {
        if (context_ != nullptr) {
            context_->AddRef();
        }
    }

    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) override {
        if (ppvObj == nullptr) {
            return E_INVALIDARG;
        }

        *ppvObj = nullptr;
        if (riid == IID_IUnknown || riid == IID_ITfEditSession) {
            *ppvObj = static_cast<ITfEditSession*>(this);
            AddRef();
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    STDMETHODIMP_(ULONG) AddRef() override {
        return static_cast<ULONG>(InterlockedIncrement(&refCount_));
    }

    STDMETHODIMP_(ULONG) Release() override {
        const LONG value = InterlockedDecrement(&refCount_);
        if (value == 0) {
            delete this;
            return 0;
        }

        return static_cast<ULONG>(value);
    }

    STDMETHODIMP DoEditSession(TfEditCookie editCookie) override {
        if (context_ == nullptr || text_.empty()) {
            return E_FAIL;
        }

        ITfInsertAtSelection* inserter = nullptr;
        HRESULT hr = context_->QueryInterface(IID_ITfInsertAtSelection, reinterpret_cast<void**>(&inserter));
        if (FAILED(hr) || inserter == nullptr) {
            return E_FAIL;
        }

        ITfRange* insertedRange = nullptr;
        hr = inserter->InsertTextAtSelection(
            editCookie,
            TF_IAS_NOQUERY,
            text_.c_str(),
            static_cast<LONG>(text_.size()),
            &insertedRange);

        if (insertedRange != nullptr) {
            insertedRange->Release();
        }

        inserter->Release();
        return hr;
    }

private:
    ~InsertTextEditSession() {
        if (context_ != nullptr) {
            context_->Release();
            context_ = nullptr;
        }
    }

    LONG refCount_;
    ITfContext* context_;
    std::wstring text_;
};

class QuerySelectionRectEditSession final : public ITfEditSession {
public:
    explicit QuerySelectionRectEditSession(ITfContext* context)
        : refCount_(1), context_(context), hasAnchor_(false) {
        if (context_ != nullptr) {
            context_->AddRef();
        }
        anchor_.x = 0;
        anchor_.y = 0;
    }

    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) override {
        if (ppvObj == nullptr) {
            return E_INVALIDARG;
        }

        *ppvObj = nullptr;
        if (riid == IID_IUnknown || riid == IID_ITfEditSession) {
            *ppvObj = static_cast<ITfEditSession*>(this);
            AddRef();
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    STDMETHODIMP_(ULONG) AddRef() override {
        return static_cast<ULONG>(InterlockedIncrement(&refCount_));
    }

    STDMETHODIMP_(ULONG) Release() override {
        const LONG value = InterlockedDecrement(&refCount_);
        if (value == 0) {
            delete this;
            return 0;
        }

        return static_cast<ULONG>(value);
    }

    STDMETHODIMP DoEditSession(TfEditCookie editCookie) override {
        if (context_ == nullptr) {
            return E_FAIL;
        }

        TF_SELECTION selection = {};
        ULONG fetched = 0;
        const HRESULT selectionHr = context_->GetSelection(editCookie, TF_DEFAULT_SELECTION, 1, &selection, &fetched);
        if (FAILED(selectionHr) || fetched == 0 || selection.range == nullptr) {
            return selectionHr;
        }

        ITfContextView* view = nullptr;
        const HRESULT viewHr = context_->GetActiveView(&view);
        if (SUCCEEDED(viewHr) && view != nullptr) {
            RECT rc = {};
            BOOL clipped = FALSE;
            const HRESULT textExtHr = view->GetTextExt(editCookie, selection.range, &rc, &clipped);
            if (SUCCEEDED(textExtHr)) {
                anchor_.x = rc.left;
                anchor_.y = rc.bottom;
                hasAnchor_ = true;
            }

            view->Release();
        }

        selection.range->Release();
        return S_OK;
    }

    bool HasAnchor() const {
        return hasAnchor_;
    }

    POINT Anchor() const {
        return anchor_;
    }

private:
    ~QuerySelectionRectEditSession() {
        if (context_ != nullptr) {
            context_->Release();
            context_ = nullptr;
        }
    }

    LONG refCount_;
    ITfContext* context_;
    bool hasAnchor_;
    POINT anchor_;
};

bool TryGetCaretScreenPointFromContext(ITfContext* context, TfClientId clientId, POINT& outPoint) {
    if (context == nullptr || clientId == TF_CLIENTID_NULL) {
        return false;
    }

    auto* session = new (std::nothrow) QuerySelectionRectEditSession(context);
    if (session == nullptr) {
        return false;
    }

    HRESULT editHr = E_FAIL;
    const HRESULT requestHr = context->RequestEditSession(
        clientId,
        session,
        TF_ES_SYNC | TF_ES_READ,
        &editHr);

    bool ok = SUCCEEDED(requestHr) && SUCCEEDED(editHr) && session->HasAnchor();
    if (ok) {
        const POINT anchor = session->Anchor();
        const int virtualLeft = GetSystemMetrics(SM_XVIRTUALSCREEN);
        const int virtualTop = GetSystemMetrics(SM_YVIRTUALSCREEN);
        const int virtualWidth = GetSystemMetrics(SM_CXVIRTUALSCREEN);
        const int virtualHeight = GetSystemMetrics(SM_CYVIRTUALSCREEN);
        const int virtualRight = virtualLeft + std::max(virtualWidth, 1);
        const int virtualBottom = virtualTop + std::max(virtualHeight, 1);
        const bool inVirtualScreen =
            anchor.x >= virtualLeft - 128 &&
            anchor.x <= virtualRight + 128 &&
            anchor.y >= virtualTop - 128 &&
            anchor.y <= virtualBottom + 128;

        if (inVirtualScreen) {
            outPoint = anchor;
        }
        else {
            ok = false;
        }
    }

    session->Release();
    return ok;
}

}  // namespace

class TextServiceClassFactory final : public IClassFactory {
public:
    TextServiceClassFactory() : refCount_(1) {
        InterlockedIncrement(&g_objectCount);
    }

    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) override {
        if (ppvObj == nullptr) {
            return E_INVALIDARG;
        }

        *ppvObj = nullptr;
        if (riid == IID_IUnknown || riid == IID_IClassFactory) {
            *ppvObj = static_cast<IClassFactory*>(this);
            AddRef();
            return S_OK;
        }

        return E_NOINTERFACE;
    }

    STDMETHODIMP_(ULONG) AddRef() override {
        return static_cast<ULONG>(InterlockedIncrement(&refCount_));
    }

    STDMETHODIMP_(ULONG) Release() override {
        const LONG value = InterlockedDecrement(&refCount_);
        if (value == 0) {
            delete this;
            return 0;
        }

        return static_cast<ULONG>(value);
    }

    STDMETHODIMP CreateInstance(IUnknown* outer, REFIID riid, void** ppvObj) override {
        if (ppvObj == nullptr) {
            return E_INVALIDARG;
        }

        *ppvObj = nullptr;
        if (outer != nullptr) {
            return CLASS_E_NOAGGREGATION;
        }

        auto* service = new (std::nothrow) TextService();
        if (service == nullptr) {
            return E_OUTOFMEMORY;
        }

        const HRESULT hr = service->QueryInterface(riid, ppvObj);
        service->Release();
        return hr;
    }

    STDMETHODIMP LockServer(BOOL lock) override {
        if (lock) {
            InterlockedIncrement(&g_dllRefCount);
        } else {
            InterlockedDecrement(&g_dllRefCount);
        }

        return S_OK;
    }

private:
    ~TextServiceClassFactory() {
        InterlockedDecrement(&g_objectCount);
    }

    LONG refCount_;
};

TextService::TextService()
    : refCount_(1),
      threadMgr_(nullptr),
      clientId_(TF_CLIENTID_NULL),
            langBarItemMgr_(nullptr),
            configLangBarItem_(nullptr),
      keyEventSinkAdvised_(false),
    threadMgrEventSinkAdvised_(false),
    threadMgrEventSinkCookie_(TF_INVALID_COOKIE),
        runtimeReady_(false),
      chineseMode_(true),
      fullShapeMode_(false),
    chinesePunctuation_(true),
    smartSymbolPairs_(true),
        autoCommitUniqueExact_(true),
        autoCommitMinCodeLength_(4),
        emptyCandidateBeep_(true),
        tabNavigation_(true),
        enterExactPriority_(true),
    contextAssociationEnabled_(true),
    contextAssociationMaxEntries_(6000),
    dictionaryProfile_(DictionaryProfile::ZhengmaLarge),
    nextSingleQuoteOpen_(true),
    nextDoubleQuoteOpen_(true),
            leftShiftTogglePending_(false),
            leftShiftToggleDownTick_(0),
        pageBoundaryDirection_(0),
        pageBoundaryHitCount_(0),
      pageSize_(kDefaultPageSize),
      toggleHotkey_(ToggleHotkey::F9),
            candidatesRevision_(0),
      pageIndex_(0),
    selectedIndexInPage_(0),
        emptyCandidateAlerted_(false),
        hasRecentAnchor_(false),
        lastAnchor_{0, 0},
                lastAnchorTick_(0),
                autoPhraseSelectedStreak_(0),
                                autoPhraseSelectedTick_(0),
                    autoPhraseDictionaryDirty_(false),
                    userFrequencyDirty_(false),
                    userDataFirstDirtyTick_(0),
                    userDataLastFlushTick_(0),
                    lastUserDataReloadCheckTick_(0),
                    lastHelperAutoPhraseReloadCheckTick_(0),
                    lastAutoPhraseBuilderLaunchTick_(0),
                    lastAutoPhraseHistoryUpdateTick_(0),
                    autoPhraseBuilderPending_(false),
                    autoPhraseBuilderProcess_(nullptr),
                                    userDataWriteStopRequested_(false),
                                    nextUserDataWriteGeneration_(0),
                    userDataStampRefreshPending_(false),
                                pageCandidatesCacheRevision_(0),
                                pageCandidatesCachePageIndex_(0),
                                pageCandidatesCachePageSize_(0),
                                    candidatesFullyExpanded_(false),
                                    deferredExpansionDueTick_(0) {
    InterlockedIncrement(&g_objectCount);
    StartRuntimeLogWorker();
}

TextService::~TextService() {
    ReapAutoPhraseBuilderProcess(false);
    if (autoPhraseBuilderProcess_ != nullptr) {
        CloseHandle(autoPhraseBuilderProcess_);
        autoPhraseBuilderProcess_ = nullptr;
    }
                    StopUserDataWriteWorker(true);
    StopRuntimeLogWorker();

    if (langBarItemMgr_ != nullptr && configLangBarItem_ != nullptr) {
        langBarItemMgr_->RemoveItem(configLangBarItem_);
    }

    if (configLangBarItem_ != nullptr) {
        configLangBarItem_->Release();
        configLangBarItem_ = nullptr;
    }

    if (langBarItemMgr_ != nullptr) {
        langBarItemMgr_->Release();
        langBarItemMgr_ = nullptr;
    }

    if (threadMgr_ != nullptr) {
        threadMgr_->Release();
        threadMgr_ = nullptr;
    }

    InterlockedDecrement(&g_objectCount);
}

STDMETHODIMP TextService::QueryInterface(REFIID riid, void** ppvObj) {
    if (ppvObj == nullptr) {
        return E_INVALIDARG;
    }

    *ppvObj = nullptr;
    if (riid == IID_IUnknown || riid == IID_ITfTextInputProcessor || riid == IID_ITfKeyEventSink || riid == IID_ITfThreadMgrEventSink) {
        if (riid == IID_ITfKeyEventSink) {
            *ppvObj = static_cast<ITfKeyEventSink*>(this);
        } else if (riid == IID_ITfThreadMgrEventSink) {
            *ppvObj = static_cast<ITfThreadMgrEventSink*>(this);
        } else {
            *ppvObj = static_cast<ITfTextInputProcessor*>(this);
        }
        AddRef();
        return S_OK;
    }

    return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) TextService::AddRef() {
    return static_cast<ULONG>(InterlockedIncrement(&refCount_));
}

STDMETHODIMP_(ULONG) TextService::Release() {
    const LONG value = InterlockedDecrement(&refCount_);
    if (value == 0) {
        delete this;
        return 0;
    }

    return static_cast<ULONG>(value);
}

STDMETHODIMP TextService::Activate(ITfThreadMgr* threadMgr, TfClientId clientId) {
    if (threadMgr == nullptr) {
        return E_INVALIDARG;
    }

    if (threadMgr_ != nullptr) {
        threadMgr_->Release();
        threadMgr_ = nullptr;
    }

    threadMgr_ = threadMgr;
    threadMgr_->AddRef();
    clientId_ = clientId;

    ITfKeystrokeMgr* keystrokeMgr = nullptr;
    const HRESULT keyMgrHr = threadMgr_->QueryInterface(IID_ITfKeystrokeMgr, reinterpret_cast<void**>(&keystrokeMgr));
    if (SUCCEEDED(keyMgrHr) && keystrokeMgr != nullptr) {
        const HRESULT adviseHr = keystrokeMgr->AdviseKeyEventSink(clientId_, static_cast<ITfKeyEventSink*>(this), TRUE);
        keyEventSinkAdvised_ = SUCCEEDED(adviseHr);
        wchar_t keySinkLog[160] = {};
        swprintf_s(keySinkLog, L"AdviseKeyEventSink hr=0x%08X advised=%d", static_cast<unsigned int>(adviseHr), keyEventSinkAdvised_ ? 1 : 0);
        Trace(keySinkLog);
        keystrokeMgr->Release();
    } else {
        keyEventSinkAdvised_ = false;
        wchar_t keyMgrLog[160] = {};
        swprintf_s(keyMgrLog, L"QueryInterface(IID_ITfKeystrokeMgr) hr=0x%08X", static_cast<unsigned int>(keyMgrHr));
        Trace(keyMgrLog);
    }

    if (langBarItemMgr_ != nullptr && configLangBarItem_ != nullptr) {
        langBarItemMgr_->RemoveItem(configLangBarItem_);
    }
    if (configLangBarItem_ != nullptr) {
        configLangBarItem_->Release();
        configLangBarItem_ = nullptr;
    }
    if (langBarItemMgr_ != nullptr) {
        langBarItemMgr_->Release();
        langBarItemMgr_ = nullptr;
    }

    const HRESULT lbMgrHr = threadMgr_->QueryInterface(IID_ITfLangBarItemMgr, reinterpret_cast<void**>(&langBarItemMgr_));
    if (SUCCEEDED(lbMgrHr) && langBarItemMgr_ != nullptr) {
        configLangBarItem_ = new (std::nothrow) ConfigLangBarButton();
        if (configLangBarItem_ != nullptr) {
            const HRESULT addHr = langBarItemMgr_->AddItem(configLangBarItem_);
            wchar_t lbLog[180] = {};
            swprintf_s(lbLog, L"LangBar AddItem hr=0x%08X", static_cast<unsigned int>(addHr));
            Trace(lbLog);
            if (FAILED(addHr)) {
                configLangBarItem_->Release();
                configLangBarItem_ = nullptr;
            }
        }
    } else {
        wchar_t lbMgrLog[160] = {};
        swprintf_s(lbMgrLog, L"QueryInterface(IID_ITfLangBarItemMgr) hr=0x%08X", static_cast<unsigned int>(lbMgrHr));
        Trace(lbMgrLog);
    }

    ITfSource* source = nullptr;
    const HRESULT sourceHr = threadMgr_->QueryInterface(IID_ITfSource, reinterpret_cast<void**>(&source));
    if (SUCCEEDED(sourceHr) && source != nullptr) {
        threadMgrEventSinkCookie_ = TF_INVALID_COOKIE;
        const HRESULT adviseHr = source->AdviseSink(
            IID_ITfThreadMgrEventSink,
            static_cast<ITfThreadMgrEventSink*>(this),
            &threadMgrEventSinkCookie_);
        threadMgrEventSinkAdvised_ = SUCCEEDED(adviseHr);
        wchar_t eventSinkLog[180] = {};
        swprintf_s(
            eventSinkLog,
            L"AdviseThreadMgrEventSink hr=0x%08X advised=%d cookie=%u",
            static_cast<unsigned int>(adviseHr),
            threadMgrEventSinkAdvised_ ? 1 : 0,
            static_cast<unsigned int>(threadMgrEventSinkCookie_));
        Trace(eventSinkLog);
        source->Release();
    } else {
        threadMgrEventSinkAdvised_ = false;
        threadMgrEventSinkCookie_ = TF_INVALID_COOKIE;
        wchar_t sourceLog[160] = {};
        swprintf_s(sourceLog, L"QueryInterface(IID_ITfSource) hr=0x%08X", static_cast<unsigned int>(sourceHr));
        Trace(sourceLog);
    }

    LoadSettings();
    if (chineseMode_) {
        EnsureRuntimeReady();
    }

    Trace(std::wstring(L"TextService activated marker=") + kBuildMarker);

    return S_OK;
}

STDMETHODIMP TextService::Deactivate() {
    FlushPendingUserDataIfNeeded(true, false);
    QueueAutoPhraseSessionWrite();
    ReapAutoPhraseBuilderProcess(false);

    if (langBarItemMgr_ != nullptr && configLangBarItem_ != nullptr) {
        langBarItemMgr_->RemoveItem(configLangBarItem_);
    }
    if (configLangBarItem_ != nullptr) {
        configLangBarItem_->Release();
        configLangBarItem_ = nullptr;
    }
    if (langBarItemMgr_ != nullptr) {
        langBarItemMgr_->Release();
        langBarItemMgr_ = nullptr;
    }

    if (threadMgr_ != nullptr && threadMgrEventSinkAdvised_ && threadMgrEventSinkCookie_ != TF_INVALID_COOKIE) {
        ITfSource* source = nullptr;
        if (SUCCEEDED(threadMgr_->QueryInterface(IID_ITfSource, reinterpret_cast<void**>(&source))) && source != nullptr) {
            source->UnadviseSink(threadMgrEventSinkCookie_);
            source->Release();
        }
    }
    threadMgrEventSinkAdvised_ = false;
    threadMgrEventSinkCookie_ = TF_INVALID_COOKIE;

    if (threadMgr_ != nullptr && keyEventSinkAdvised_) {
        ITfKeystrokeMgr* keystrokeMgr = nullptr;
        if (SUCCEEDED(threadMgr_->QueryInterface(IID_ITfKeystrokeMgr, reinterpret_cast<void**>(&keystrokeMgr))) && keystrokeMgr != nullptr) {
            keystrokeMgr->UnadviseKeyEventSink(clientId_);
            keystrokeMgr->Release();
        }
        keyEventSinkAdvised_ = false;
    }

    if (threadMgr_ != nullptr) {
        threadMgr_->Release();
        threadMgr_ = nullptr;
    }

    clientId_ = TF_CLIENTID_NULL;
    ClearComposition();
    candidateWindow_.Hide();
    Trace(L"TextService deactivated");
    return S_OK;
}

void TextService::MarkAutoPhraseDictionaryDirty() {
    if (!autoPhraseDictionaryDirty_ && !userFrequencyDirty_) {
        userDataFirstDirtyTick_ = GetTickCount64();
    }
    autoPhraseDictionaryDirty_ = true;
}

void TextService::MarkFrequencyDataDirty() {
    if (!userFrequencyDirty_ && !autoPhraseDictionaryDirty_) {
        userDataFirstDirtyTick_ = GetTickCount64();
    }
    userFrequencyDirty_ = true;
}

void TextService::SyncUserDataFilesStamp() {
    userDataFilesStamp_ = BuildUserDataFilesStamp(
        userDictPath_,
        autoPhraseDictPath_,
        userFreqPath_,
        blockedEntriesPath_,
        contextAssocPath_,
        contextAssocBlacklistPath_);
}

bool TextService::ReloadUserDataIfChanged(bool force) {
    if (userDataDir_.empty()) {
        return false;
    }

    if (!force) {
        const ULONGLONG now = GetTickCount64();
        if (lastUserDataReloadCheckTick_ != 0 &&
            now >= lastUserDataReloadCheckTick_ &&
            (now - lastUserDataReloadCheckTick_) < kUserDataReloadMinIntervalMs) {
            return false;
        }
        lastUserDataReloadCheckTick_ = now;
    } else {
        lastUserDataReloadCheckTick_ = GetTickCount64();
    }

    if (userDataStampRefreshPending_ && !HasPendingTrackedUserDataWrites()) {
        SyncUserDataFilesStamp();
        userDataStampRefreshPending_ = false;
    }

    const std::wstring currentStamp = BuildUserDataFilesStamp(
        userDictPath_,
        autoPhraseDictPath_,
        userFreqPath_,
        blockedEntriesPath_,
        contextAssocPath_,
        contextAssocBlacklistPath_);
    if (!force && currentStamp == userDataFilesStamp_) {
        return false;
    }

    const bool reloaded = ReloadActiveDictionaries();
    Trace(L"user data reloaded changed=" + std::to_wstring(force ? 1 : 0) + L" success=" + std::to_wstring(reloaded ? 1 : 0));
    return reloaded;
}

bool TextService::ReloadHelperAutoPhraseEntriesIfChanged(bool force) {
    if (autoPhraseHelperPath_.empty()) {
        helperAutoPhraseEntriesByCode_.clear();
        helperAutoPhraseStamp_.clear();
        return false;
    }

    if (!force) {
        const ULONGLONG now = GetTickCount64();
        if (lastHelperAutoPhraseReloadCheckTick_ != 0 &&
            now >= lastHelperAutoPhraseReloadCheckTick_ &&
            (now - lastHelperAutoPhraseReloadCheckTick_) < kHelperAutoPhraseReloadMinIntervalMs) {
            return false;
        }
        lastHelperAutoPhraseReloadCheckTick_ = now;
    } else {
        lastHelperAutoPhraseReloadCheckTick_ = GetTickCount64();
    }

    const std::wstring currentStamp = QueryFileStampToken(autoPhraseHelperPath_);
    if (!force && currentStamp == helperAutoPhraseStamp_) {
        return false;
    }

    const bool loaded = LoadHelperAutoPhraseEntries();
    helperAutoPhraseStamp_ = currentStamp;
    Trace(L"helper auto phrase reloaded force=" + std::to_wstring(force ? 1 : 0) + L" loaded=" + std::to_wstring(loaded ? 1 : 0));
    return true;
}

bool TextService::EnsureRuntimeReady() {
    if (runtimeReady_) {
        return true;
    }

    if (EnsureUserDataDirectory(userDataDir_)) {
        MigrateLegacyUserDataFiles(userDataDir_);
        userDictPath_ = BuildUserDataFilePath(userDataDir_, L"yuninput_user.dict");
        autoPhraseUserPath_ = BuildUserDataFilePath(userDataDir_, L"auto_phrase_runtime.dict");
        autoPhraseHelperPath_ = BuildUserDataFilePath(userDataDir_, L"auto_phrase_helper.dict");
        userFreqPath_ = BuildUserDataFilePath(userDataDir_, L"user_freq.txt");
        blockedEntriesPath_ = BuildUserDataFilePath(userDataDir_, L"blocked_entries.txt");
        contextAssocPath_ = BuildUserDataFilePath(userDataDir_, L"context_assoc.txt");
        contextAssocBlacklistPath_ = BuildUserDataFilePath(userDataDir_, L"context_assoc_blacklist.txt");
        manualPhraseReviewPath_ = BuildUserDataFilePath(userDataDir_, L"manual_phrase_review.txt");
        autoPhraseSessionPath_ = BuildUserDataFilePath(userDataDir_, L"auto_phrase_session.txt");
    }

    StartUserDataWriteWorker();
    const bool dictionariesLoaded = LoadConfiguredDictionaries();

    if (!userDataDir_.empty()) {
        engine_.LoadFrequencyFromFile(userFreqPath_);
        engine_.LoadBlockedEntriesFromFile(blockedEntriesPath_);
        LoadContextAssociationFromFile(contextAssocPath_);
        LoadContextAssociationBlacklistFromFile(contextAssocBlacklistPath_);
        LoadAutoPhraseSessionState();
        autoPhraseBuilderPending_ = true;
        LaunchAutoPhraseBuilderProcess();
        if (FileContainsLegacyFrequencyNoise(userFreqPath_)) {
            engine_.SaveFrequencyToFile(userFreqPath_);
        }
        SyncUserDataFilesStamp();
        helperAutoPhraseStamp_ = QueryFileStampToken(autoPhraseHelperPath_);
    }

    candidateWindow_.EnsureCreated();
    candidateWindow_.SetAsyncPollCallback([this]() {
        RunDeferredCandidateExpansion();
    });

    runtimeReady_ = true;
    Trace(L"runtime init completed dict=" + std::to_wstring(dictionariesLoaded ? 1 : 0));
    return dictionariesLoaded;
}

void TextService::FlushPendingUserDataIfNeeded(bool force, bool waitForCompletion) {
    if (!autoPhraseDictionaryDirty_ && !userFrequencyDirty_) {
        if (force && waitForCompletion) {
            WaitForUserDataWrites(false);
            SyncUserDataFilesStamp();
        }
        return;
    }

    const ULONGLONG now = GetTickCount64();
    const ULONGLONG requiredInterval =
        autoPhraseDictionaryDirty_ && userFrequencyDirty_
            ? std::min(kAutoPhraseFlushIntervalMs, kUserFrequencyFlushIntervalMs)
            : (autoPhraseDictionaryDirty_ ? kAutoPhraseFlushIntervalMs : kUserFrequencyFlushIntervalMs);
    if (!force) {
        if (userDataFirstDirtyTick_ == 0) {
            userDataFirstDirtyTick_ = now;
        }

        const ULONGLONG elapsed = now - userDataFirstDirtyTick_;
        if (elapsed < requiredInterval) {
            return;
        }
    }

    bool queuedAny = false;
    if (autoPhraseDictionaryDirty_ && !userDictPath_.empty()) {
        const std::string userSnapshot = engine_.BuildUserDictionaryFileContent();
        {
            std::lock_guard<std::mutex> lock(userDataWriteMutex_);
            pendingUserDictWrite_.path = userDictPath_;
            pendingUserDictWrite_.content = userSnapshot;
            pendingUserDictWrite_.deleteIfEmpty = false;
            pendingUserDictWrite_.generation = ++nextUserDataWriteGeneration_;

            pendingAutoPhraseDictWrite_.path = autoPhraseUserPath_;
            pendingAutoPhraseDictWrite_.content = engine_.BuildAutoPhraseDictionaryFileContent(kRoamingAutoPhraseMaxEntries);
            pendingAutoPhraseDictWrite_.deleteIfEmpty = true;
            pendingAutoPhraseDictWrite_.generation = ++nextUserDataWriteGeneration_;
        }
        userDataWriteCv_.notify_all();
        autoPhraseDictionaryDirty_ = false;
        queuedAny = true;
        userDataStampRefreshPending_ = true;
    }

    if (userFrequencyDirty_ && !userFreqPath_.empty()) {
        const std::string snapshot = engine_.BuildFrequencyFileContent();
        {
            std::lock_guard<std::mutex> lock(userDataWriteMutex_);
            pendingUserFreqWrite_.path = userFreqPath_;
            pendingUserFreqWrite_.content = snapshot;
            pendingUserFreqWrite_.deleteIfEmpty = false;
            pendingUserFreqWrite_.generation = ++nextUserDataWriteGeneration_;
        }
        userDataWriteCv_.notify_all();
        userFrequencyDirty_ = false;
        queuedAny = true;
        userDataStampRefreshPending_ = true;
    }

    if (queuedAny) {
        userDataFirstDirtyTick_ = 0;
        userDataLastFlushTick_ = now;
    }

    if (force && waitForCompletion) {
        WaitForUserDataWrites(false);
        if (queuedAny) {
            SyncUserDataFilesStamp();
            userDataStampRefreshPending_ = false;
        }
    }
}

void TextService::StartUserDataWriteWorker() {
    if (userDataWriteThread_.joinable()) {
        return;
    }

    userDataWriteStopRequested_ = false;
    userDataWriteThread_ = std::thread(&TextService::UserDataWriteWorkerMain, this);
}

void TextService::StopUserDataWriteWorker(bool waitForPending) {
    if (!userDataWriteThread_.joinable()) {
        return;
    }

    if (waitForPending) {
        WaitForUserDataWrites(true);
    }

    {
        std::lock_guard<std::mutex> lock(userDataWriteMutex_);
        userDataWriteStopRequested_ = true;
    }
    userDataWriteCv_.notify_all();
    userDataWriteThread_.join();
}

void TextService::UserDataWriteWorkerMain() {
    for (;;) {
        PendingAsyncFileWrite userDictWrite;
        PendingAsyncFileWrite autoPhraseDictWrite;
        PendingAsyncFileWrite userFreqWrite;
        PendingAsyncFileWrite contextAssocWrite;
        PendingAsyncFileWrite blockedEntriesWrite;
        PendingAsyncFileWrite autoPhraseSessionWrite;
        PendingAsyncFileWrite manualPhraseReviewWrite;

        {
            std::unique_lock<std::mutex> lock(userDataWriteMutex_);
            userDataWriteCv_.wait(lock, [this]() {
                return userDataWriteStopRequested_ ||
                    pendingUserDictWrite_.generation > pendingUserDictWrite_.completedGeneration ||
                    pendingAutoPhraseDictWrite_.generation > pendingAutoPhraseDictWrite_.completedGeneration ||
                    pendingUserFreqWrite_.generation > pendingUserFreqWrite_.completedGeneration ||
                    pendingContextAssocWrite_.generation > pendingContextAssocWrite_.completedGeneration ||
                    pendingBlockedEntriesWrite_.generation > pendingBlockedEntriesWrite_.completedGeneration ||
                    pendingAutoPhraseSessionWrite_.generation > pendingAutoPhraseSessionWrite_.completedGeneration ||
                    pendingManualPhraseReviewWrite_.generation > pendingManualPhraseReviewWrite_.completedGeneration;
            });

            if (userDataWriteStopRequested_ &&
                pendingUserDictWrite_.generation == pendingUserDictWrite_.completedGeneration &&
                pendingAutoPhraseDictWrite_.generation == pendingAutoPhraseDictWrite_.completedGeneration &&
                pendingUserFreqWrite_.generation == pendingUserFreqWrite_.completedGeneration &&
                pendingContextAssocWrite_.generation == pendingContextAssocWrite_.completedGeneration &&
                pendingBlockedEntriesWrite_.generation == pendingBlockedEntriesWrite_.completedGeneration &&
                pendingAutoPhraseSessionWrite_.generation == pendingAutoPhraseSessionWrite_.completedGeneration &&
                pendingManualPhraseReviewWrite_.generation == pendingManualPhraseReviewWrite_.completedGeneration) {
                break;
            }

            if (pendingUserDictWrite_.generation > pendingUserDictWrite_.completedGeneration) {
                userDictWrite = pendingUserDictWrite_;
            }
            if (pendingAutoPhraseDictWrite_.generation > pendingAutoPhraseDictWrite_.completedGeneration) {
                autoPhraseDictWrite = pendingAutoPhraseDictWrite_;
            }
            if (pendingUserFreqWrite_.generation > pendingUserFreqWrite_.completedGeneration) {
                userFreqWrite = pendingUserFreqWrite_;
            }
            if (pendingContextAssocWrite_.generation > pendingContextAssocWrite_.completedGeneration) {
                contextAssocWrite = pendingContextAssocWrite_;
            }
            if (pendingBlockedEntriesWrite_.generation > pendingBlockedEntriesWrite_.completedGeneration) {
                blockedEntriesWrite = pendingBlockedEntriesWrite_;
            }
            if (pendingAutoPhraseSessionWrite_.generation > pendingAutoPhraseSessionWrite_.completedGeneration) {
                autoPhraseSessionWrite = pendingAutoPhraseSessionWrite_;
            }
            if (pendingManualPhraseReviewWrite_.generation > pendingManualPhraseReviewWrite_.completedGeneration) {
                manualPhraseReviewWrite = pendingManualPhraseReviewWrite_;
                pendingManualPhraseReviewWrite_.content.clear();
            }
        }

        if (userDictWrite.generation != 0) {
            WriteUtf8FileSnapshot(userDictWrite.path, userDictWrite.content, userDictWrite.deleteIfEmpty, userDictWrite.append);
            std::lock_guard<std::mutex> lock(userDataWriteMutex_);
            pendingUserDictWrite_.completedGeneration = std::max(pendingUserDictWrite_.completedGeneration, userDictWrite.generation);
            userDataWriteCv_.notify_all();
        }

        if (autoPhraseDictWrite.generation != 0) {
            WriteUtf8FileSnapshot(autoPhraseDictWrite.path, autoPhraseDictWrite.content, autoPhraseDictWrite.deleteIfEmpty, autoPhraseDictWrite.append);
            std::lock_guard<std::mutex> lock(userDataWriteMutex_);
            pendingAutoPhraseDictWrite_.completedGeneration = std::max(pendingAutoPhraseDictWrite_.completedGeneration, autoPhraseDictWrite.generation);
            userDataWriteCv_.notify_all();
        }

        if (userFreqWrite.generation != 0) {
            WriteUtf8FileSnapshot(userFreqWrite.path, userFreqWrite.content, userFreqWrite.deleteIfEmpty, userFreqWrite.append);
            std::lock_guard<std::mutex> lock(userDataWriteMutex_);
            pendingUserFreqWrite_.completedGeneration = std::max(pendingUserFreqWrite_.completedGeneration, userFreqWrite.generation);
            userDataWriteCv_.notify_all();
        }

        if (contextAssocWrite.generation != 0) {
            WriteUtf8FileSnapshot(contextAssocWrite.path, contextAssocWrite.content, contextAssocWrite.deleteIfEmpty, contextAssocWrite.append);
            std::lock_guard<std::mutex> lock(userDataWriteMutex_);
            pendingContextAssocWrite_.completedGeneration = std::max(pendingContextAssocWrite_.completedGeneration, contextAssocWrite.generation);
            userDataWriteCv_.notify_all();
        }

        if (blockedEntriesWrite.generation != 0) {
            WriteUtf8FileSnapshot(blockedEntriesWrite.path, blockedEntriesWrite.content, blockedEntriesWrite.deleteIfEmpty, blockedEntriesWrite.append);
            std::lock_guard<std::mutex> lock(userDataWriteMutex_);
            pendingBlockedEntriesWrite_.completedGeneration = std::max(pendingBlockedEntriesWrite_.completedGeneration, blockedEntriesWrite.generation);
            userDataWriteCv_.notify_all();
        }

        if (autoPhraseSessionWrite.generation != 0) {
            WriteUtf8FileSnapshot(autoPhraseSessionWrite.path, autoPhraseSessionWrite.content, autoPhraseSessionWrite.deleteIfEmpty, autoPhraseSessionWrite.append);
            std::lock_guard<std::mutex> lock(userDataWriteMutex_);
            pendingAutoPhraseSessionWrite_.completedGeneration = std::max(pendingAutoPhraseSessionWrite_.completedGeneration, autoPhraseSessionWrite.generation);
            userDataWriteCv_.notify_all();
        }

        if (manualPhraseReviewWrite.generation != 0) {
            WriteUtf8FileSnapshot(manualPhraseReviewWrite.path, manualPhraseReviewWrite.content, manualPhraseReviewWrite.deleteIfEmpty, manualPhraseReviewWrite.append);
            std::lock_guard<std::mutex> lock(userDataWriteMutex_);
            pendingManualPhraseReviewWrite_.completedGeneration = std::max(pendingManualPhraseReviewWrite_.completedGeneration, manualPhraseReviewWrite.generation);
            userDataWriteCv_.notify_all();
        }
    }
}

void TextService::QueueAutoPhraseSessionWrite() {
    if (autoPhraseSessionPath_.empty()) {
        return;
    }

    bool deleteFile = false;
    const std::string snapshot = BuildAutoPhraseSessionStateSnapshot(deleteFile);
    if (!WriteUtf8FileSnapshot(autoPhraseSessionPath_, snapshot, deleteFile, false)) {
        return;
    }

    autoPhraseSessionStamp_ = QueryFileStampToken(autoPhraseSessionPath_);
    if (deleteFile) {
        autoPhraseBuilderPending_ = false;
        lastAutoPhraseHistoryUpdateTick_ = 0;
        return;
    }

    autoPhraseBuilderPending_ = true;
    lastAutoPhraseHistoryUpdateTick_ = GetTickCount64();
    LaunchAutoPhraseBuilderProcess();
}

void TextService::QueueManualPhraseReviewAppend(const std::string& line) {
    if (manualPhraseReviewPath_.empty() || line.empty()) {
        return;
    }

    {
        std::lock_guard<std::mutex> lock(userDataWriteMutex_);
        pendingManualPhraseReviewWrite_.path = manualPhraseReviewPath_;
        pendingManualPhraseReviewWrite_.content.append(line);
        pendingManualPhraseReviewWrite_.deleteIfEmpty = false;
        pendingManualPhraseReviewWrite_.append = true;
        pendingManualPhraseReviewWrite_.generation = ++nextUserDataWriteGeneration_;
    }
    userDataWriteCv_.notify_all();
}

bool TextService::WaitForUserDataWrites(bool includeSessionWrite) {
    if (!userDataWriteThread_.joinable()) {
        return true;
    }

    std::unique_lock<std::mutex> lock(userDataWriteMutex_);
    const std::uint64_t targetUserDictGeneration = pendingUserDictWrite_.generation;
    const std::uint64_t targetAutoPhraseDictGeneration = pendingAutoPhraseDictWrite_.generation;
    const std::uint64_t targetUserFreqGeneration = pendingUserFreqWrite_.generation;
    const std::uint64_t targetContextAssocGeneration = pendingContextAssocWrite_.generation;
    const std::uint64_t targetBlockedEntriesGeneration = pendingBlockedEntriesWrite_.generation;
    const std::uint64_t targetSessionGeneration = pendingAutoPhraseSessionWrite_.generation;
    const std::uint64_t targetManualPhraseReviewGeneration = pendingManualPhraseReviewWrite_.generation;

    userDataWriteCv_.wait(lock, [this, includeSessionWrite, targetUserDictGeneration, targetAutoPhraseDictGeneration, targetUserFreqGeneration, targetContextAssocGeneration, targetBlockedEntriesGeneration, targetSessionGeneration, targetManualPhraseReviewGeneration]() {
        const bool userDictDone = pendingUserDictWrite_.completedGeneration >= targetUserDictGeneration;
        const bool autoPhraseDictDone = pendingAutoPhraseDictWrite_.completedGeneration >= targetAutoPhraseDictGeneration;
        const bool userFreqDone = pendingUserFreqWrite_.completedGeneration >= targetUserFreqGeneration;
        const bool contextAssocDone = pendingContextAssocWrite_.completedGeneration >= targetContextAssocGeneration;
        const bool blockedEntriesDone = pendingBlockedEntriesWrite_.completedGeneration >= targetBlockedEntriesGeneration;
        const bool sessionDone = pendingAutoPhraseSessionWrite_.completedGeneration >= targetSessionGeneration;
        const bool manualPhraseReviewDone = pendingManualPhraseReviewWrite_.completedGeneration >= targetManualPhraseReviewGeneration;
        return userDictDone && autoPhraseDictDone && userFreqDone && contextAssocDone && blockedEntriesDone && manualPhraseReviewDone && (!includeSessionWrite || sessionDone);
    });

    return true;
}

bool TextService::HasPendingTrackedUserDataWrites() {
    std::lock_guard<std::mutex> lock(userDataWriteMutex_);
    return pendingUserDictWrite_.generation > pendingUserDictWrite_.completedGeneration ||
           pendingAutoPhraseDictWrite_.generation > pendingAutoPhraseDictWrite_.completedGeneration ||
           pendingUserFreqWrite_.generation > pendingUserFreqWrite_.completedGeneration ||
           pendingContextAssocWrite_.generation > pendingContextAssocWrite_.completedGeneration ||
           pendingBlockedEntriesWrite_.generation > pendingBlockedEntriesWrite_.completedGeneration;
}

bool TextService::WriteUtf8FileSnapshot(const std::wstring& filePath, const std::string& content, bool deleteIfEmpty, bool append) {
    if (filePath.empty()) {
        return false;
    }

    const std::filesystem::path path(filePath);
    std::error_code ec;
    if (deleteIfEmpty && content.empty()) {
        std::filesystem::remove(path, ec);
        return !ec;
    }

    if (path.has_parent_path()) {
        std::filesystem::create_directories(path.parent_path(), ec);
        ec.clear();
    }

    const std::ios::openmode mode = std::ios::out | std::ios::binary | (append ? std::ios::app : std::ios::trunc);
    std::ofstream output(path, mode);
    if (!output) {
        return false;
    }

    output.write(content.data(), static_cast<std::streamsize>(content.size()));
    return output.good();
}

std::string TextService::BuildAutoPhraseSessionStateSnapshot(bool& deleteFile) const {
    deleteFile = autoPhraseHistoryText_.empty();
    if (deleteFile) {
        return std::string();
    }

    std::ostringstream output;
    output << "boot_tick\t" << GetTickCount64() << '\n';
    output << "history\t" << WideToUtf8Text(autoPhraseHistoryText_) << '\n';
    return output.str();
}

bool TextService::ReloadAutoPhraseSessionStateIfChanged(bool force) {
    if (autoPhraseSessionPath_.empty()) {
        return false;
    }

    const std::wstring currentStamp = QueryFileStampToken(autoPhraseSessionPath_);
    if (!force && currentStamp == autoPhraseSessionStamp_) {
        return false;
    }

    return LoadAutoPhraseSessionState();
}

bool TextService::ReapAutoPhraseBuilderProcess(bool reloadSessionState) {
    if (autoPhraseBuilderProcess_ == nullptr) {
        return false;
    }

    const DWORD waitResult = WaitForSingleObject(autoPhraseBuilderProcess_, 0);
    if (waitResult == WAIT_TIMEOUT) {
        return false;
    }

    DWORD exitCode = STILL_ACTIVE;
    if (!GetExitCodeProcess(autoPhraseBuilderProcess_, &exitCode)) {
        exitCode = GetLastError();
        Trace(L"auto phrase builder exit-code query failed error=" + std::to_wstring(exitCode));
    } else {
        Trace(L"auto phrase builder completed exit=" + std::to_wstring(exitCode));
    }

    CloseHandle(autoPhraseBuilderProcess_);
    autoPhraseBuilderProcess_ = nullptr;

    if (reloadSessionState) {
        ReloadAutoPhraseSessionStateIfChanged(true);
    }

    return true;
}

bool TextService::LaunchAutoPhraseBuilderProcess() {
    ReapAutoPhraseBuilderProcess(true);

    if (autoPhraseSessionPath_.empty() || autoPhraseHistoryText_.empty()) {
        autoPhraseBuilderPending_ = false;
        return false;
    }

    if (autoPhraseBuilderProcess_ != nullptr) {
        if (SignalAutoPhraseHelperWakeEvent()) {
            return true;
        }

        Trace(L"auto phrase builder still running; skip relaunch");
        return false;
    }

    if (!autoPhraseBuilderPending_) {
        return SignalAutoPhraseHelperWakeEvent();
    }

    if (SignalAutoPhraseHelperWakeEvent()) {
        autoPhraseBuilderPending_ = false;
        return true;
    }

    const ULONGLONG now = GetTickCount64();
    if (lastAutoPhraseHistoryUpdateTick_ != 0 &&
        lastAutoPhraseBuilderLaunchTick_ != 0 &&
        now >= lastAutoPhraseHistoryUpdateTick_ &&
        (now - lastAutoPhraseHistoryUpdateTick_) < kAutoPhraseBuilderLaunchMinIntervalMs) {
        return false;
    }

    std::wstring modulePath;
    if (!GetModulePath(modulePath)) {
        return false;
    }

    if (autoPhraseBuilderPath_.empty()) {
        const std::filesystem::path moduleFile(modulePath);
        const std::filesystem::path moduleDir = moduleFile.parent_path();
        const std::filesystem::path rootDir = ResolveDataRootFromModule(modulePath);
        const std::filesystem::path releaseDir = rootDir / L"build" / L"Release";
        const std::filesystem::path binDir = rootDir / L"bin";
        const std::filesystem::path candidates[] = {
            moduleDir / L"yuninput_user_dict_builder.exe",
            releaseDir / L"yuninput_user_dict_builder.exe",
            binDir / L"yuninput_user_dict_builder.exe",
        };

        for (const auto& candidate : candidates) {
            std::error_code ec;
            if (std::filesystem::exists(candidate, ec)) {
                autoPhraseBuilderPath_ = candidate.wstring();
                break;
            }
        }
    }

    if (autoPhraseBuilderPath_.empty()) {
        Trace(L"auto phrase builder executable missing");
        return false;
    }

    const std::filesystem::path rootDir = ResolveDataRootFromModule(modulePath);
    const std::filesystem::path dataDir = rootDir / L"data";
    const std::wstring profileDictPath = dictionaryProfile_ == DictionaryProfile::ZhengmaLargePinyin
        ? (dataDir / L"zhengma-pinyin.dict").wstring()
        : (dataDir / L"zhengma-all.dict").wstring();
    const std::wstring fallbackDictPath = (dataDir / L"yuninput_basic.dict").wstring();
    const std::wstring phraseSourcePath = std::filesystem::exists(dataDir / L"zhengma-single.dict")
        ? (dataDir / L"zhengma-single.dict").wstring()
        : profileDictPath;

    std::wstring commandLine = QuoteCommandLineArg(autoPhraseBuilderPath_) +
        L" session-watch " +
        QuoteCommandLineArg(autoPhraseSessionPath_) + L" " +
        QuoteCommandLineArg(profileDictPath) + L" " +
        QuoteCommandLineArg(fallbackDictPath) + L" " +
        QuoteCommandLineArg(phraseSourcePath) + L" " +
        QuoteCommandLineArg(userDictPath_) + L" " +
        QuoteCommandLineArg(autoPhraseDictPath_) + L" " +
        QuoteCommandLineArg(autoPhraseUserPath_) + L" " +
        QuoteCommandLineArg(autoPhraseHelperPath_);

    Trace(L"auto phrase builder launch command=" + commandLine);

    STARTUPINFOW startupInfo = {};
    startupInfo.cb = sizeof(startupInfo);
    PROCESS_INFORMATION processInfo = {};
    std::vector<wchar_t> mutableCommand(commandLine.begin(), commandLine.end());
    mutableCommand.push_back(L'\0');
    if (!CreateProcessW(
            nullptr,
            mutableCommand.data(),
            nullptr,
            nullptr,
            FALSE,
            CREATE_NO_WINDOW,
            nullptr,
            std::filesystem::path(autoPhraseBuilderPath_).parent_path().c_str(),
            &startupInfo,
            &processInfo)) {
        const DWORD error = GetLastError();
        Trace(L"auto phrase builder launch failed error=" + std::to_wstring(error));
        return false;
    }

    CloseHandle(processInfo.hThread);
    autoPhraseBuilderProcess_ = processInfo.hProcess;
    lastAutoPhraseBuilderLaunchTick_ = now;
    autoPhraseBuilderPending_ = false;
    Trace(L"auto phrase builder launched pid=" + std::to_wstring(static_cast<unsigned long long>(processInfo.dwProcessId)));
    return true;
}

std::string TextService::BuildContextAssociationFileContent() const {
    std::ostringstream output;

    for (const auto& pair : contextAssociationScores_) {
        const size_t split = pair.first.find(L'\t');
        if (split == std::wstring::npos) {
            continue;
        }

        const std::wstring prevText = pair.first.substr(0, split);
        const std::wstring nextText = pair.first.substr(split + 1);
        output << WideToUtf8Text(prevText) << ' ' << WideToUtf8Text(nextText) << ' ' << pair.second << '\n';
    }

    return output.str();
}

bool TextService::SaveAutoPhraseSessionState() const {
    if (autoPhraseSessionPath_.empty()) {
        return false;
    }

    bool deleteFile = false;
    const std::string snapshot = BuildAutoPhraseSessionStateSnapshot(deleteFile);
    return WriteUtf8FileSnapshot(autoPhraseSessionPath_, snapshot, deleteFile, false);
}

bool TextService::LoadAutoPhraseSessionState() {
    autoPhraseHistoryText_.clear();
    sessionAutoPhraseEntries_.clear();

    if (autoPhraseSessionPath_.empty()) {
        autoPhraseSessionStamp_.clear();
        return false;
    }

    std::ifstream input(autoPhraseSessionPath_, std::ios::in | std::ios::binary);
    if (!input) {
        autoPhraseSessionStamp_ = QueryFileStampToken(autoPhraseSessionPath_);
        return false;
    }

    const ULONGLONG currentBootTick = GetTickCount64();
    bool rebootInvalidated = false;
    std::string line;
    while (std::getline(input, line)) {
        if (line.empty()) {
            continue;
        }

        std::istringstream iss(line);
        std::string tag;
        if (!std::getline(iss, tag, '\t')) {
            continue;
        }

        if (tag == "boot_tick") {
            std::string tickText;
            if (std::getline(iss, tickText, '\t')) {
                const ULONGLONG savedBootTick = static_cast<ULONGLONG>(_strtoui64(tickText.c_str(), nullptr, 10));
                if (savedBootTick > currentBootTick) {
                    rebootInvalidated = true;
                    break;
                }
            }
            continue;
        }

        if (tag == "history") {
            std::string historyUtf8;
            if (std::getline(iss, historyUtf8, '\t')) {
                autoPhraseHistoryText_ = Utf8ToWideText(historyUtf8);
                if (autoPhraseHistoryText_.size() > kSessionAutoPhraseHistoryCharLimit) {
                    autoPhraseHistoryText_.erase(0, autoPhraseHistoryText_.size() - kSessionAutoPhraseHistoryCharLimit);
                }
            }
            continue;
        }

        if (tag == "entry") {
            std::string textUtf8;
            std::string codesCsv;
            std::string tickText;
            if (!std::getline(iss, textUtf8, '\t') ||
                !std::getline(iss, codesCsv, '\t') ||
                !std::getline(iss, tickText, '\t')) {
                continue;
            }

            SessionAutoPhraseEntry entry;
            entry.text = Utf8ToWideText(textUtf8);
            entry.lastTick = static_cast<ULONGLONG>(_strtoui64(tickText.c_str(), nullptr, 10));
            std::istringstream codeStream(codesCsv);
            std::string codeUtf8;
            while (std::getline(codeStream, codeUtf8, ',')) {
                const std::wstring code = Utf8ToWideText(codeUtf8);
                if (code.empty()) {
                    continue;
                }
                if (!engine_.HasEntry(code, entry.text)) {
                    entry.codes.push_back(code);
                }
            }

            if (!entry.text.empty() && !entry.codes.empty()) {
                sessionAutoPhraseEntries_[entry.text] = std::move(entry);
            }
        }
    }

    if (rebootInvalidated) {
        autoPhraseHistoryText_.clear();
        sessionAutoPhraseEntries_.clear();
        QueueAutoPhraseSessionWrite();
        autoPhraseSessionStamp_ = QueryFileStampToken(autoPhraseSessionPath_);
        return false;
    }

    if (!autoPhraseHistoryText_.empty()) {
        CollectSessionAutoPhraseCandidatesForTail(GetTickCount64());
    }

    if (PruneSessionAutoPhraseEntries()) {
        QueueAutoPhraseSessionWrite();
    }

    autoPhraseSessionStamp_ = QueryFileStampToken(autoPhraseSessionPath_);
    return !autoPhraseHistoryText_.empty() || !sessionAutoPhraseEntries_.empty();
}

bool TextService::LoadHelperAutoPhraseEntries() {
    helperAutoPhraseEntriesByCode_.clear();

    if (autoPhraseHelperPath_.empty()) {
        return false;
    }

    std::ifstream input(autoPhraseHelperPath_, std::ios::in | std::ios::binary);
    if (!input) {
        return false;
    }

    std::string line;
    while (std::getline(input, line)) {
        if (line.empty() || line[0] == '#') {
            continue;
        }

        std::istringstream iss(line);
        std::string firstToken;
        std::string secondToken;
        if (!(iss >> firstToken >> secondToken)) {
            continue;
        }

        const bool firstIsCode = IsAsciiCodeToken(firstToken);
        const bool secondIsCode = IsAsciiCodeToken(secondToken);
        if (firstIsCode == secondIsCode) {
            continue;
        }

        const std::string codeUtf8 = firstIsCode ? firstToken : secondToken;
        const std::string textUtf8 = firstIsCode ? secondToken : firstToken;

        std::wstring code = Utf8ToWideText(codeUtf8);
        std::transform(
            code.begin(),
            code.end(),
            code.begin(),
            [](wchar_t ch) {
                return static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch)));
            });
        code.erase(
            std::remove_if(
                code.begin(),
                code.end(),
                [](wchar_t ch) {
                    return ch < L'a' || ch > L'z';
                }),
            code.end());

        const std::wstring text = Utf8ToWideText(textUtf8);
        if (code.empty() || text.empty()) {
            continue;
        }

        auto& entries = helperAutoPhraseEntriesByCode_[code];
        const auto duplicateIt = std::find_if(
            entries.begin(),
            entries.end(),
            [&text](const HelperAutoPhraseEntry& entry) {
                return entry.text == text;
            });
        if (duplicateIt != entries.end()) {
            continue;
        }

        HelperAutoPhraseEntry entry;
        entry.code = code;
        entry.text = text;
        entries.push_back(std::move(entry));
    }

    return !helperAutoPhraseEntriesByCode_.empty();
}

bool TextService::IsHanCharacter(wchar_t ch) {
    return (ch >= 0x3400 && ch <= 0x4DBF) ||
           (ch >= 0x4E00 && ch <= 0x9FFF) ||
           (ch >= 0xF900 && ch <= 0xFAFF);
}

void TextService::RecordSessionAutoPhraseBreak() {
    if (!autoPhraseHistoryText_.empty() && autoPhraseHistoryText_.back() != kSessionAutoPhraseBreak) {
        autoPhraseHistoryText_.push_back(kSessionAutoPhraseBreak);
        if (autoPhraseHistoryText_.size() > kSessionAutoPhraseHistoryCharLimit) {
            autoPhraseHistoryText_.erase(0, autoPhraseHistoryText_.size() - kSessionAutoPhraseHistoryCharLimit);
        }
        PruneSessionAutoPhraseEntries();
        QueueAutoPhraseSessionWrite();
    }
}

void TextService::CollectSessionAutoPhraseCandidatesForTail(ULONGLONG now) {
    if (autoPhraseHistoryText_.size() < 2) {
        return;
    }

    const auto sanitizePhraseCode = [](std::wstring code) {
        code.erase(
            std::remove_if(
                code.begin(),
                code.end(),
                [](wchar_t ch) {
                    return ch < L'a' || ch > L'z';
                }),
            code.end());
        if (code.size() > 20) {
            code.resize(20);
        }
        return code;
    };

    size_t tailStart = autoPhraseHistoryText_.find_last_of(kSessionAutoPhraseBreak);
    if (tailStart == std::wstring::npos) {
        tailStart = 0;
    } else {
        tailStart += 1;
    }

    if (tailStart >= autoPhraseHistoryText_.size()) {
        return;
    }

    const std::wstring tailText = autoPhraseHistoryText_.substr(tailStart);
    if (tailText.size() < 2) {
        return;
    }

    const size_t maxPhraseLength = std::min(kSessionAutoPhraseMaxLength, tailText.size());
    for (size_t start = 0; start + 2 <= tailText.size(); ++start) {
        const size_t remaining = tailText.size() - start;
        const size_t maxLengthForStart = std::min(maxPhraseLength, remaining);
        for (size_t phraseLength = 2; phraseLength <= maxLengthForStart; ++phraseLength) {
            const std::wstring phraseText = tailText.substr(start, phraseLength);
            auto existingIt = sessionAutoPhraseEntries_.find(phraseText);
            if (existingIt != sessionAutoPhraseEntries_.end()) {
                existingIt->second.lastTick = now;
                continue;
            }

            std::vector<std::wstring> phraseCodes;
            if (!phraseBuildEngine_.TryBuildPhraseCodes(phraseText, phraseCodes) || phraseCodes.empty()) {
                continue;
            }

            std::vector<std::wstring> sanitizedCodes;
            sanitizedCodes.reserve(phraseCodes.size());
            for (std::wstring& phraseCode : phraseCodes) {
                phraseCode = sanitizePhraseCode(std::move(phraseCode));
                if (phraseCode.empty()) {
                    continue;
                }
                if (std::find(sanitizedCodes.begin(), sanitizedCodes.end(), phraseCode) != sanitizedCodes.end()) {
                    continue;
                }
                if (engine_.HasEntry(phraseCode, phraseText)) {
                    continue;
                }
                sanitizedCodes.push_back(std::move(phraseCode));
            }

            if (sanitizedCodes.empty()) {
                continue;
            }

            SessionAutoPhraseEntry entry;
            entry.text = phraseText;
            entry.codes = std::move(sanitizedCodes);
            entry.lastTick = now;
            sessionAutoPhraseEntries_.emplace(entry.text, std::move(entry));
        }
    }
}

bool TextService::PruneSessionAutoPhraseEntries() {
    if (sessionAutoPhraseEntries_.empty()) {
        return false;
    }

    std::unordered_set<std::wstring> validPhrases;
    validPhrases.reserve(sessionAutoPhraseEntries_.size() * 2);

    size_t segmentStart = 0;
    while (segmentStart < autoPhraseHistoryText_.size()) {
        while (segmentStart < autoPhraseHistoryText_.size() && autoPhraseHistoryText_[segmentStart] == kSessionAutoPhraseBreak) {
            ++segmentStart;
        }
        if (segmentStart >= autoPhraseHistoryText_.size()) {
            break;
        }

        size_t segmentEnd = segmentStart;
        while (segmentEnd < autoPhraseHistoryText_.size() && autoPhraseHistoryText_[segmentEnd] != kSessionAutoPhraseBreak) {
            ++segmentEnd;
        }

        const size_t segmentLength = segmentEnd - segmentStart;
        if (segmentLength >= 2) {
            const size_t maxPhraseLength = std::min(kSessionAutoPhraseMaxLength, segmentLength);
            for (size_t start = 0; start + 2 <= segmentLength; ++start) {
                const size_t remaining = segmentLength - start;
                const size_t maxLengthForStart = std::min(maxPhraseLength, remaining);
                for (size_t phraseLength = 2; phraseLength <= maxLengthForStart; ++phraseLength) {
                    validPhrases.insert(autoPhraseHistoryText_.substr(segmentStart + start, phraseLength));
                }
            }
        }

        segmentStart = segmentEnd + 1;
    }

    bool changed = false;
    for (auto it = sessionAutoPhraseEntries_.begin(); it != sessionAutoPhraseEntries_.end();) {
        if (it->second.codes.empty() || validPhrases.find(it->first) == validPhrases.end()) {
            it = sessionAutoPhraseEntries_.erase(it);
            changed = true;
        } else {
            ++it;
        }
    }

    return changed;
}

void TextService::UpdateSessionAutoPhraseHistory(const std::wstring& committedText, ULONGLONG now) {
    bool changed = false;
    for (wchar_t ch : committedText) {
        if (IsHanCharacter(ch)) {
            autoPhraseHistoryText_.push_back(ch);
            if (autoPhraseHistoryText_.size() > kSessionAutoPhraseHistoryCharLimit) {
                autoPhraseHistoryText_.erase(0, autoPhraseHistoryText_.size() - kSessionAutoPhraseHistoryCharLimit);
            }
            changed = true;
        } else if (!autoPhraseHistoryText_.empty() && autoPhraseHistoryText_.back() != kSessionAutoPhraseBreak) {
            autoPhraseHistoryText_.push_back(kSessionAutoPhraseBreak);
            if (autoPhraseHistoryText_.size() > kSessionAutoPhraseHistoryCharLimit) {
                autoPhraseHistoryText_.erase(0, autoPhraseHistoryText_.size() - kSessionAutoPhraseHistoryCharLimit);
            }
            changed = true;
        }
    }

    if (changed) {
        CollectSessionAutoPhraseCandidatesForTail(now);
        PruneSessionAutoPhraseEntries();
        QueueAutoPhraseSessionWrite();
    }
}

void TextService::MergeSessionAutoPhraseCandidates() {
    if (compositionCode_.size() < 4 || sessionAutoPhraseEntries_.empty()) {
        return;
    }

    std::vector<CandidateItem> sessionCandidates;
    sessionCandidates.reserve(16);
    for (const auto& pair : sessionAutoPhraseEntries_) {
        const SessionAutoPhraseEntry& entry = pair.second;
        if (entry.text.empty()) {
            continue;
        }

        if (std::find(entry.codes.begin(), entry.codes.end(), compositionCode_) == entry.codes.end()) {
            continue;
        }

        CandidateItem candidate;
        candidate.text = entry.text;
        candidate.code = compositionCode_;
        candidate.commitCode = compositionCode_;
        candidate.exactMatch = true;
        candidate.fromSessionAutoPhrase = true;
        candidate.consumedLength = compositionCode_.size();
        sessionCandidates.push_back(std::move(candidate));
    }

    for (CandidateItem& candidate : sessionCandidates) {
        allCandidates_.push_back(std::move(candidate));
    }
}

bool TextService::PromoteSessionAutoPhrase(const std::wstring& text) {
    const auto it = sessionAutoPhraseEntries_.find(text);
    if (it == sessionAutoPhraseEntries_.end()) {
        return false;
    }

    SessionAutoPhraseEntry entry = it->second;
    entry.lastTick = GetTickCount64();

    bool changed = false;
    for (const std::wstring& code : entry.codes) {
        if (code.empty()) {
            continue;
        }

        changed = engine_.AddUserEntry(code, entry.text) || changed;
        AppendPhraseReviewEntry(code, entry.text, L"auto-commit");
    }

    if (changed) {
        MarkAutoPhraseDictionaryDirty();
    }

    sessionAutoPhraseEntries_.erase(it);
    QueueAutoPhraseSessionWrite();
    return changed;
}

void TextService::RefreshCandidates(bool expandAll) {
    const LONGLONG refreshStartCounter = QueryPerfCounterValue();
    allCandidates_.clear();
    InvalidatePageCandidatesCache();
    ReloadUserDataIfChanged(false);
    ReloadHelperAutoPhraseEntriesIfChanged(false);
    candidatesFullyExpanded_ = expandAll;
    if (compositionCode_.empty()) {
        pageIndex_ = 0;
        selectedIndexInPage_ = 0;
        emptyCandidateAlerted_ = false;
        UpdateCandidateWindow();
        return;
    }

    const size_t fastTargetCandidateCount = std::max<size_t>(pageSize_ + 2, std::min<size_t>(pageSize_ * 2, 12));
    const size_t targetCandidateCount = expandAll ? kRefreshTargetCandidateCount : fastTargetCandidateCount;
    const size_t primaryQueryLimit = expandAll ? kPrimaryCandidateQueryLimit : std::max<size_t>(pageSize_ * 2, 12);

    const bool pinyinFallbackMode = dictionaryProfile_ == DictionaryProfile::ZhengmaLargePinyin;
    const bool allowIncrementalPrefixCache = !expandAll && !pinyinFallbackMode && compositionCode_.size() > 1;

    auto buildDisplayCode = [this, pinyinFallbackMode](const std::wstring& text, const std::wstring& rawCode) {
        if (!pinyinFallbackMode) {
            return rawCode;
        }

        const std::wstring zhengmaHint = GetSingleCharZhengmaCodeHint(text);
        if (!zhengmaHint.empty()) {
            return zhengmaHint;
        }

        return std::wstring(L"-");
    };

    std::unordered_map<std::wstring, size_t> candidateIndexByKey;
    candidateIndexByKey.reserve(primaryQueryLimit + 64);

    auto makeCandidateMergeKey = [](const std::wstring& text, const std::wstring& commitCode) {
        std::wstring key;
        key.reserve(text.size() + 1 + commitCode.size());
        key.append(text);
        key.push_back(L'\t');
        key.append(commitCode);
        return key;
    };

    auto mergeCandidate = [this, &candidateIndexByKey, &makeCandidateMergeKey](
                              const std::wstring& text,
                              const std::wstring& displayCode,
                              const std::wstring& commitCode,
                              std::uint64_t contextScore,
                              std::uint64_t learnedScore,
                              bool exactMatch,
                              bool boostedUser,
                              bool boostedLearned,
                              bool boostedContext,
                              bool fromAutoPhrase,
                              bool fromSessionAutoPhrase,
                              bool fromSystemDict,
                              size_t consumedLength) {
        if (text.empty()) {
            return;
        }

        const std::wstring stableCommitCode = commitCode.empty() ? displayCode : commitCode;
        const std::wstring candidateKey = makeCandidateMergeKey(text, stableCommitCode);

        const auto existingIt = candidateIndexByKey.find(candidateKey);
        if (existingIt != candidateIndexByKey.end()) {
            CandidateItem& existing = allCandidates_[existingIt->second];

            existing.boostedUser = existing.boostedUser || boostedUser;
            existing.boostedLearned = existing.boostedLearned || boostedLearned;
            existing.boostedContext = existing.boostedContext || boostedContext;
            existing.exactMatch = existing.exactMatch || exactMatch;
            existing.fromAutoPhrase = existing.fromAutoPhrase || fromAutoPhrase;
            existing.fromSessionAutoPhrase = existing.fromSessionAutoPhrase || fromSessionAutoPhrase;
            existing.fromSystemDict = existing.fromSystemDict || fromSystemDict;
            if (learnedScore > existing.learnedScore) {
                existing.learnedScore = learnedScore;
            }
            if (contextScore > existing.contextScore) {
                existing.contextScore = contextScore;
            }

            const bool preferNewCode =
                existing.code.empty() ||
                (!displayCode.empty() && displayCode.size() < existing.code.size()) ||
                (boostedContext && !existing.boostedContext);
            if (preferNewCode) {
                existing.code = displayCode;
            }

            if (existing.commitCode.empty()) {
                existing.commitCode = stableCommitCode;
            }
            if (consumedLength > existing.consumedLength) {
                existing.consumedLength = consumedLength;
            }
            return;
        }

        CandidateItem candidate;
        candidate.text = text;
        candidate.code = displayCode;
        candidate.commitCode = stableCommitCode;
        candidate.contextScore = contextScore;
        candidate.learnedScore = learnedScore;
        candidate.exactMatch = exactMatch;
        candidate.boostedUser = boostedUser;
        candidate.boostedLearned = boostedLearned;
        candidate.boostedContext = boostedContext;
        candidate.fromAutoPhrase = fromAutoPhrase;
        candidate.fromSessionAutoPhrase = fromSessionAutoPhrase;
        candidate.fromSystemDict = fromSystemDict;
        candidate.consumedLength = consumedLength;
        allCandidates_.push_back(std::move(candidate));
        candidateIndexByKey.emplace(candidateKey, allCandidates_.size() - 1);
    };

    std::vector<CandidateItem> helperTailCandidates;

    if (allowIncrementalPrefixCache) {
        const std::wstring previousCode = compositionCode_.substr(0, compositionCode_.size() - 1);
        const auto cacheIt = compositionCandidateCache_.find(previousCode);
        if (cacheIt != compositionCandidateCache_.end()) {
            const auto& cachedCandidates = cacheIt->second.candidates;
            allCandidates_.reserve(std::max(allCandidates_.capacity(), cachedCandidates.size() + 32));
            for (const CandidateItem& cached : cachedCandidates) {
                std::wstring stableCode = cached.commitCode.empty() ? cached.code : cached.commitCode;
                if (stableCode.empty()) {
                    continue;
                }

                if (stableCode.size() < compositionCode_.size() ||
                    stableCode.compare(0, compositionCode_.size(), compositionCode_) != 0) {
                    continue;
                }

                CandidateItem candidate = cached;
                candidate.exactMatch = stableCode == compositionCode_;
                candidate.consumedLength = compositionCode_.size();
                if (candidate.code.empty()) {
                    candidate.code = stableCode;
                }

                const std::wstring candidateKey = makeCandidateMergeKey(candidate.text, stableCode);
                if (candidateIndexByKey.find(candidateKey) != candidateIndexByKey.end()) {
                    continue;
                }

                allCandidates_.push_back(std::move(candidate));
                candidateIndexByKey.emplace(candidateKey, allCandidates_.size() - 1);
            }
        }
    }

    bool hasExactCurrentCode = std::any_of(
        allCandidates_.begin(),
        allCandidates_.end(),
        [](const CandidateItem& item) {
            return item.exactMatch;
        });
    const bool canUseIncrementalPrefixFastPath =
        allowIncrementalPrefixCache &&
        hasExactCurrentCode &&
        allCandidates_.size() >= pageSize_;

    const size_t fastQueryScanBudget =
        expandAll
            ? 0
            : (compositionCode_.size() <= 1
                   ? kFastQueryScanBudgetOneChar
                   : (compositionCode_.size() == 2 ? kFastQueryScanBudgetTwoChar : 0));

    const std::vector<CompositionEngine::Entry> queried =
        canUseIncrementalPrefixFastPath
            ? std::vector<CompositionEngine::Entry>()
            : (fastQueryScanBudget == 0
                   ? engine_.QueryCandidateEntries(compositionCode_, primaryQueryLimit)
                   : engine_.QueryCandidateEntriesFast(compositionCode_, primaryQueryLimit, fastQueryScanBudget));
    if (!canUseIncrementalPrefixFastPath) {
        allCandidates_.reserve(std::max(allCandidates_.capacity(), queried.size() + 32));
    }

    for (const auto& item : queried) {
        if (pinyinFallbackMode && item.text.size() != 1) {
            continue;
        }
        if (pinyinFallbackMode && item.code != compositionCode_) {
            continue;
        }
        hasExactCurrentCode = hasExactCurrentCode || item.code == compositionCode_;
        const std::wstring displayCode = buildDisplayCode(item.text, item.code);

        mergeCandidate(
            item.text,
            displayCode,
            item.code.empty() ? compositionCode_ : item.code,
            0,
            item.learnedScore,
            item.code == compositionCode_,
            item.isUser && !item.isAutoPhrase,
            item.isLearned,
            false,
            item.isAutoPhrase,
            false,
            !item.isUser,
            compositionCode_.size());
    }

    auto needsMoreCandidates = [this, targetCandidateCount]() {
        return allCandidates_.size() < targetCandidateCount;
    };

    const bool fastPathSatisfied = !expandAll && hasExactCurrentCode && allCandidates_.size() >= pageSize_;

    if (!fastPathSatisfied && pinyinFallbackMode && allCandidates_.empty() && compositionCode_.size() > 2) {
        std::wstring prefixCode = compositionCode_;
        for (size_t prefixLength = compositionCode_.size() - 1; prefixLength >= 2; --prefixLength) {
            prefixCode.resize(prefixLength);
            const std::vector<CompositionEngine::Entry> prefixMatches = engine_.QueryCandidateEntries(prefixCode, kPrefixCandidateQueryLimit);
            for (const auto& item : prefixMatches) {
                if (item.text.size() != 1 || item.code != prefixCode) {
                    continue;
                }

                const std::wstring displayCode = buildDisplayCode(item.text, item.code);

                mergeCandidate(
                    item.text,
                    displayCode,
                    prefixCode,
                    0,
                    item.learnedScore,
                    false,
                    item.isUser && !item.isAutoPhrase,
                    item.isLearned,
                    false,
                    item.isAutoPhrase,
                    false,
                    !item.isUser,
                    prefixLength);
            }

            if (!allCandidates_.empty()) {
                break;
            }
            if (prefixLength == 2) {
                break;
            }
        }
    }

    if (!fastPathSatisfied && !hasExactCurrentCode && compositionCode_.size() > 1 && !pinyinFallbackMode && needsMoreCandidates()) {
        const size_t prefixFallbackTarget = std::max(targetCandidateCount, pageSize_ * 3);
        std::wstring prefixCode = compositionCode_;
        for (size_t prefixLength = compositionCode_.size() - 1; prefixLength >= 1; --prefixLength) {
            prefixCode.resize(prefixLength);
            const std::vector<CompositionEngine::Entry> prefixMatches = engine_.QueryCandidateEntries(prefixCode, kPrefixCandidateQueryLimit);
            for (const auto& item : prefixMatches) {
                if (item.code != prefixCode) {
                    continue;
                }

                mergeCandidate(
                    item.text,
                    item.code,
                    prefixCode,
                    0,
                    item.learnedScore,
                    false,
                    item.isUser && !item.isAutoPhrase,
                    item.isLearned,
                    false,
                    item.isAutoPhrase,
                    false,
                    !item.isUser,
                    prefixLength);
            }

            if (allCandidates_.size() >= prefixFallbackTarget) {
                break;
            }

            if (prefixLength == 1) {
                break;
            }
        }
    }

    const ULONGLONG now = GetTickCount64();
    if (!fastPathSatisfied && !recentCommits_.empty() && !pinyinFallbackMode && needsMoreCandidates()) {
        const CommitHistoryItem& previous = recentCommits_.back();
        if ((now - previous.tick) <= 8000ULL && !previous.code.empty() && !previous.text.empty()) {
            const std::wstring phraseCode = previous.code + compositionCode_;
            const std::vector<CompositionEngine::Entry> phraseMatches = engine_.QueryCandidateEntries(phraseCode, kPhraseCandidateQueryLimit);
            for (const auto& item : phraseMatches) {
                if (item.text.size() <= previous.text.size()) {
                    continue;
                }
                if (item.text.compare(0, previous.text.size(), previous.text) != 0) {
                    continue;
                }

                const std::wstring suffixText = item.text.substr(previous.text.size());
                if (suffixText.empty() || suffixText.size() > 4) {
                    continue;
                }

                mergeCandidate(
                    suffixText,
                    compositionCode_,
                    compositionCode_,
                    0,
                    item.learnedScore,
                    false,
                    item.isUser && !item.isAutoPhrase,
                    true,
                    true,
                    item.isAutoPhrase,
                    false,
                    !item.isUser,
                    compositionCode_.size());
            }
        }
    }

    if (!fastPathSatisfied && contextAssociationEnabled_ && !recentCommits_.empty() && !allCandidates_.empty()) {
        const CommitHistoryItem& previous = recentCommits_.back();
        if ((now - previous.tick) <= 15000ULL && !previous.text.empty()) {
            const std::wstring associationPrefix = previous.text + L"\t";
            std::wstring associationKey = associationPrefix;
            associationKey.reserve(associationPrefix.size() + 32);
            for (CandidateItem& candidate : allCandidates_) {
                associationKey.resize(associationPrefix.size());
                associationKey.append(candidate.text);
                if (contextAssociationBlacklist_.find(associationKey) != contextAssociationBlacklist_.end()) {
                    continue;
                }

                const auto scoreIt = contextAssociationScores_.find(associationKey);
                if (scoreIt == contextAssociationScores_.end()) {
                    continue;
                }

                const std::uint64_t contextScore = scoreIt->second;
                if (contextScore == 0) {
                    continue;
                }

                candidate.contextScore = std::max(candidate.contextScore, contextScore);
                candidate.boostedContext = true;
            }
        }
    }

    if (!fastPathSatisfied && !pinyinFallbackMode && compositionCode_.size() >= 4) {
        for (const auto& pair : sessionAutoPhraseEntries_) {
            const SessionAutoPhraseEntry& entry = pair.second;
            if (entry.text.empty()) {
                continue;
            }
            if (std::find(entry.codes.begin(), entry.codes.end(), compositionCode_) == entry.codes.end()) {
                continue;
            }

            mergeCandidate(
                entry.text,
                compositionCode_,
                compositionCode_,
                0,
                0,
                true,
                false,
                false,
                false,
                false,
                true,
                false,
                compositionCode_.size());
        }

        const auto helperIt = helperAutoPhraseEntriesByCode_.find(compositionCode_);
        if (helperIt != helperAutoPhraseEntriesByCode_.end()) {
            helperTailCandidates.reserve(helperIt->second.size());
            for (const HelperAutoPhraseEntry& entry : helperIt->second) {
                if (entry.text.empty()) {
                    continue;
                }

                const std::wstring candidateKey = makeCandidateMergeKey(entry.text, compositionCode_);
                if (candidateIndexByKey.find(candidateKey) != candidateIndexByKey.end()) {
                    continue;
                }

                CandidateItem candidate;
                candidate.text = entry.text;
                candidate.code = compositionCode_;
                candidate.commitCode = compositionCode_;
                candidate.exactMatch = true;
                candidate.fromAutoPhrase = true;
                candidate.consumedLength = compositionCode_.size();
                helperTailCandidates.push_back(std::move(candidate));
            }
        }
    }

    const bool repeatBoostActive =
        autoPhraseSelectedStreak_ > 3 &&
        !lastAutoPhraseSelectedKey_.empty() &&
        now >= autoPhraseSelectedTick_ &&
        (now - autoPhraseSelectedTick_) <= 60000ULL;

    auto fillSortHints = [this](CandidateItem& candidate) {
        candidate.sortAutoOnly = candidate.fromAutoPhrase && !candidate.fromSystemDict && !candidate.boostedUser;
        candidate.sortSingleChar = candidate.text.size() == 1;
        candidate.sortGB2312Text = IsTextInGB2312Cached(candidate.text);
        candidate.sortNonGB2312Single = candidate.sortSingleChar && !candidate.sortGB2312Text;

        size_t sortCodeLength = candidate.code.empty() ? candidate.commitCode.size() : candidate.code.size();
        if (sortCodeLength == 0) {
            sortCodeLength = compositionCode_.size();
        }

        const bool oneCodeSingle =
            candidate.sortSingleChar && candidate.code.size() <= 1 && candidate.commitCode.size() <= 1;
        candidate.sortOneCodeSingle = oneCodeSingle;
        candidate.sortOneCodeSingleUsed = oneCodeSingle && (candidate.boostedUser || candidate.boostedLearned);

        const bool exactTwoCodePhrase =
            candidate.text.size() == 2 &&
            candidate.code.size() == 2 &&
            candidate.commitCode.size() == 2;
        const bool twoCodeSingle =
            candidate.sortSingleChar && candidate.code.size() == 2 && candidate.commitCode.size() == 2;
        candidate.sortTwoCodeSingleOrPhrase = twoCodeSingle || exactTwoCodePhrase;
        candidate.sortSystemFiveCodePhrase = candidate.fromSystemDict && candidate.text.size() >= 2 && sortCodeLength == 5;

        if (candidate.sortOneCodeSingleUsed) {
            candidate.sortShortCodeTier = 0;
        } else if (candidate.sortOneCodeSingle) {
            candidate.sortShortCodeTier = 1;
        } else if (candidate.sortTwoCodeSingleOrPhrase) {
            candidate.sortShortCodeTier = 2;
        } else {
            candidate.sortShortCodeTier = 3;
        }

        candidate.sortPreferredPhrase =
            candidate.text.size() >= 2 &&
            candidate.exactMatch &&
            sortCodeLength >= 4 &&
            (candidate.fromSessionAutoPhrase ||
             candidate.fromAutoPhrase ||
             candidate.boostedUser ||
             candidate.boostedLearned ||
             candidate.boostedContext ||
             candidate.learnedScore > 0 ||
             candidate.contextScore > 0);

        std::uint8_t tier = 7;
        if (candidate.sortNonGB2312Single) {
            tier = 250;
        } else if (candidate.sortSystemFiveCodePhrase) {
            tier = 240;
        } else if (candidate.sortPreferredPhrase) {
            tier = 0;
        } else if (candidate.sortSingleChar && sortCodeLength <= 1) {
            tier = 1;
        } else if (candidate.text.size() == 2 && sortCodeLength == 2) {
            tier = 2;
        } else if (candidate.sortSingleChar && sortCodeLength == 2) {
            tier = 3;
        } else if (candidate.sortSingleChar && sortCodeLength == 3) {
            tier = 4;
        } else if (!candidate.exactMatch && sortCodeLength >= 4) {
            tier = 5;
        } else if (candidate.fromAutoPhrase && sortCodeLength == 4) {
            tier = 6;
        } else if (!candidate.fromAutoPhrase && candidate.text.size() >= 2 && sortCodeLength == 4) {
            tier = 7;
        } else if (candidate.sortSingleChar && sortCodeLength == 4) {
            tier = 8;
        }
        candidate.sortPrimaryTier = tier;
    };

    for (CandidateItem& candidate : allCandidates_) {
        candidate.boostedAutoRepeat = false;
        if (repeatBoostActive && candidate.fromAutoPhrase) {
            const std::wstring candidateKey =
                (candidate.commitCode.empty() ? compositionCode_ : candidate.commitCode) + L"\t" + candidate.text;
            candidate.boostedAutoRepeat = candidateKey == lastAutoPhraseSelectedKey_;
        }
        fillSortHints(candidate);
    }

    const size_t queryCodeLength = compositionCode_.size();
    std::stable_sort(
        allCandidates_.begin(),
        allCandidates_.end(),
        [pinyinFallbackMode, queryCodeLength](const CandidateItem& left, const CandidateItem& right) {
            if (pinyinFallbackMode) {
                if (left.sortSingleChar != right.sortSingleChar) {
                    return left.sortSingleChar > right.sortSingleChar;
                }
            }
            if (left.exactMatch != right.exactMatch) {
                return left.exactMatch > right.exactMatch;
            }
            if (left.boostedUser != right.boostedUser) {
                return left.boostedUser > right.boostedUser;
            }
            if (left.boostedLearned != right.boostedLearned) {
                return left.boostedLearned > right.boostedLearned;
            }
            if (left.learnedScore != right.learnedScore) {
                return left.learnedScore > right.learnedScore;
            }
            if (left.fromSessionAutoPhrase != right.fromSessionAutoPhrase) {
                return left.fromSessionAutoPhrase > right.fromSessionAutoPhrase;
            }
            if (left.sortPreferredPhrase != right.sortPreferredPhrase) {
                return left.sortPreferredPhrase > right.sortPreferredPhrase;
            }
            if (left.sortPrimaryTier != right.sortPrimaryTier) {
                return left.sortPrimaryTier < right.sortPrimaryTier;
            }
            if (queryCodeLength <= 2) {
                if (left.sortShortCodeTier != right.sortShortCodeTier) {
                    return left.sortShortCodeTier < right.sortShortCodeTier;
                }
            }
            if (left.sortGB2312Text != right.sortGB2312Text) {
                return left.sortGB2312Text > right.sortGB2312Text;
            }
            if (left.contextScore != right.contextScore) {
                return left.contextScore > right.contextScore;
            }
            if (left.boostedContext != right.boostedContext) {
                return left.boostedContext > right.boostedContext;
            }
            if (left.boostedAutoRepeat != right.boostedAutoRepeat) {
                return left.boostedAutoRepeat > right.boostedAutoRepeat;
            }
            if (left.sortSystemFiveCodePhrase != right.sortSystemFiveCodePhrase) {
                return left.sortSystemFiveCodePhrase < right.sortSystemFiveCodePhrase;
            }
            if (left.sortNonGB2312Single != right.sortNonGB2312Single) {
                return left.sortNonGB2312Single < right.sortNonGB2312Single;
            }
            if (left.sortAutoOnly != right.sortAutoOnly) {
                return left.sortAutoOnly < right.sortAutoOnly;
            }
            if (left.consumedLength != right.consumedLength) {
                return left.consumedLength > right.consumedLength;
            }
            return false;
        });

    for (CandidateItem& candidate : helperTailCandidates) {
        allCandidates_.push_back(std::move(candidate));
    }

    pageIndex_ = 0;
    selectedIndexInPage_ = FindPreferredSelectionIndexForPage(pageIndex_);
    if (allCandidates_.empty()) {
        if (emptyCandidateBeep_ && !emptyCandidateAlerted_) {
            MessageBeep(MB_ICONWARNING);
            emptyCandidateAlerted_ = true;
        }
    } else {
        emptyCandidateAlerted_ = false;
    }

    CacheCurrentCandidatesForCode();
    UpdateCandidateWindow();
    if (!expandAll) {
        ScheduleDeferredCandidateExpansion();
    }
    TraceLatencySample(
        L"refresh code=" + compositionCode_ +
            L" expand=" + std::to_wstring(expandAll ? 1 : 0) +
            L" count=" + std::to_wstring(allCandidates_.size()),
        refreshStartCounter,
        QueryPerfCounterValue());
}

const wchar_t* TextService::GetDictionaryProfileName(DictionaryProfile profile) {
    switch (profile) {
    case DictionaryProfile::ZhengmaLargePinyin:
        return L"zhengma-large-pinyin";
    case DictionaryProfile::ZhengmaLarge:
    default:
        return L"zhengma-all";
    }
}

bool TextService::ReloadActiveDictionaries() {
    FlushPendingUserDataIfNeeded(true);
    const bool loaded = LoadConfiguredDictionaries();

    if (!userFreqPath_.empty()) {
        engine_.LoadFrequencyFromFile(userFreqPath_);
    }
    if (!blockedEntriesPath_.empty()) {
        engine_.LoadBlockedEntriesFromFile(blockedEntriesPath_);
    }
    if (!contextAssocPath_.empty()) {
        LoadContextAssociationFromFile(contextAssocPath_);
    }
    if (!contextAssocBlacklistPath_.empty()) {
        LoadContextAssociationBlacklistFromFile(contextAssocBlacklistPath_);
    }

    SyncUserDataFilesStamp();

    return loaded;
}

bool TextService::SwitchDictionaryProfile(DictionaryProfile profile) {
    if (dictionaryProfile_ == profile) {
        Trace(std::wstring(L"dictionary profile unchanged=") + GetDictionaryProfileName(profile));
        NotifyDictionaryProfileSwitch(profile, true);
        return true;
    }

    const DictionaryProfile previousProfile = dictionaryProfile_;
    dictionaryProfile_ = profile;
    ClearComposition();
    candidateWindow_.Hide();

    bool loaded = ReloadActiveDictionaries();
    if (!loaded) {
        dictionaryProfile_ = previousProfile;
        ReloadActiveDictionaries();
    }

    const bool persisted = loaded ? PersistDictionaryProfileSetting() : false;
    Trace(
        std::wstring(L"dictionary profile switched=") + GetDictionaryProfileName(dictionaryProfile_) +
        L" loaded=" + (loaded ? L"1" : L"0") +
        L" persisted=" + (persisted ? L"1" : L"0"));

    NotifyDictionaryProfileSwitch(dictionaryProfile_, loaded);
    return loaded;
}

bool TextService::LoadConfiguredDictionaries() {
    std::wstring modulePath;
    if (!GetModulePath(modulePath)) {
        return false;
    }

    const auto root = ResolveDataRootFromModule(modulePath);
    const auto dataDir = root / "data";
    const auto packagedUserDictPath = dataDir / "yuninput_user.dict";
    const auto packagedAutoPhraseDictPath = dataDir / "yuninput_user-extend.dict";
    const auto singleCharSourcePath = dataDir / "zhengma-single.dict";
    autoPhraseDictPath_ = packagedAutoPhraseDictPath.wstring();

    if (userDictPath_.empty()) {
        userDictPath_ = packagedUserDictPath.wstring();
    } else {
        CopyFileIfMissing(packagedUserDictPath, std::filesystem::path(userDictPath_));
    }

    std::filesystem::path profileDictPath;
    std::wstring profileName = GetDictionaryProfileName(dictionaryProfile_);
    switch (dictionaryProfile_) {
    case DictionaryProfile::ZhengmaLargePinyin:
        profileDictPath = dataDir / "zhengma-pinyin.dict";
        break;
    case DictionaryProfile::ZhengmaLarge:
    default:
        profileDictPath = dataDir / "zhengma-all.dict";
        break;
    }

    bool loadedProfileDict = false;
    if (std::filesystem::exists(profileDictPath)) {
        loadedProfileDict = engine_.LoadDictionaryFromFile(profileDictPath.wstring());
    }

    bool loadedFallback = false;
    if (!loadedProfileDict) {
        const auto fallbackDictPath = dataDir / "yuninput_basic.dict";
        if (std::filesystem::exists(fallbackDictPath)) {
            loadedFallback = engine_.LoadDictionaryFromFile(fallbackDictPath.wstring());
        }
    }

    phraseBuildEngine_ = CompositionEngine{};
    bool loadedPhraseBuildSource = false;
    if (std::filesystem::exists(singleCharSourcePath)) {
        loadedPhraseBuildSource = phraseBuildEngine_.LoadDictionaryFromFile(singleCharSourcePath.wstring());
    }
    if (!loadedPhraseBuildSource && loadedProfileDict) {
        loadedPhraseBuildSource = phraseBuildEngine_.LoadDictionaryFromFile(profileDictPath.wstring());
    }
    phraseBuildEngine_.LoadDictionaryMetadataOnlyFromFile(packagedUserDictPath.wstring());

    if (loadedPhraseBuildSource) {
        singleCharZhengmaCodeHints_ = phraseBuildEngine_.BuildSingleCharCodeHintMap();
    }
    if (singleCharZhengmaCodeHints_.empty()) {
        EnsureSingleCharZhengmaCodeHintsLoaded(dataDir);
    }

    const bool loadedPackagedUser = engine_.LoadUserDictionaryFromFile(userDictPath_);
    const bool loadedPackagedAutoPhrase = engine_.LoadAutoPhraseDictionaryFromFile(autoPhraseDictPath_);
    const bool loadedRoamingAutoPhrase = !autoPhraseUserPath_.empty() && engine_.LoadAutoPhraseDictionaryFromFile(autoPhraseUserPath_);
    const bool loadedHelperAutoPhrase = LoadHelperAutoPhraseEntries();
        helperAutoPhraseStamp_ = QueryFileStampToken(autoPhraseHelperPath_);

    Trace(
        L"Dictionary profile=" + profileName +
        L" profile_loaded=" + (loadedProfileDict ? L"1" : L"0") +
        L" fallback_loaded=" + (loadedFallback ? L"1" : L"0") +
        L" packaged_user=" + (loadedPackagedUser ? L"1" : L"0") +
        L" packaged_extend=" + (loadedPackagedAutoPhrase ? L"1" : L"0") +
        L" roaming_auto=" + (loadedRoamingAutoPhrase ? L"1" : L"0") +
        L" helper_auto=" + (loadedHelperAutoPhrase ? L"1" : L"0") +
        L" phrase_source=" + (loadedPhraseBuildSource ? L"1" : L"0") +
        (dictionaryProfile_ == DictionaryProfile::ZhengmaLargePinyin ? L" mode=pinyin-single-char-with-zhengma-hint" : L""));

    return loadedProfileDict || loadedFallback;
}

void TextService::EnsureSingleCharZhengmaCodeHintsLoaded(const std::filesystem::path& dataDir) {
    singleCharZhengmaCodeHints_.clear();

    std::filesystem::path zhengmaPath = dataDir / "zhengma-single.dict";
    if (!std::filesystem::exists(zhengmaPath)) {
        zhengmaPath = dataDir / "zhengma-large.dict";
    }
    std::ifstream input(zhengmaPath, std::ios::in | std::ios::binary);
    if (!input) {
        Trace(L"load zhengma hint map failed path=" + zhengmaPath.wstring());
        return;
    }

    std::string line;
    while (std::getline(input, line)) {
        if (line.empty() || line[0] == '#') {
            continue;
        }

        std::istringstream iss(line);
        std::string codeUtf8;
        std::string textUtf8;
        if (!(iss >> codeUtf8 >> textUtf8)) {
            continue;
        }

        std::wstring code = Utf8ToWideText(codeUtf8);
        std::transform(
            code.begin(),
            code.end(),
            code.begin(),
            [](wchar_t ch) {
                return static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch)));
            });
        code.erase(
            std::remove_if(
                code.begin(),
                code.end(),
                [](wchar_t ch) {
                    return ch < L'a' || ch > L'z';
                }),
            code.end());
        const std::wstring text = Utf8ToWideText(textUtf8);
        if (text.size() != 1 || code.empty() || code.size() > 4) {
            continue;
        }

        const wchar_t ch = text[0];
        auto it = singleCharZhengmaCodeHints_.find(ch);
        if (it == singleCharZhengmaCodeHints_.end()) {
            singleCharZhengmaCodeHints_[ch] = code;
            continue;
        }

        const size_t existingLen = it->second.size();
        const size_t candidateLen = code.size();
        const bool existingIsFullCode = existingLen == 4;
        const bool candidateIsFullCode = candidateLen == 4;

        // For hint display, prefer full 4-code Zhengma; if absent, keep the longest available code.
        const bool shouldReplace =
            (candidateIsFullCode && !existingIsFullCode) ||
            (!existingIsFullCode && !candidateIsFullCode && candidateLen > existingLen);
        if (shouldReplace) {
            it->second = code;
        }
    }

    Trace(L"load zhengma hint map count=" + std::to_wstring(singleCharZhengmaCodeHints_.size()));
}

std::wstring TextService::GetSingleCharZhengmaCodeHint(const std::wstring& text) const {
    if (text.size() != 1) {
        return L"";
    }

    const auto it = singleCharZhengmaCodeHints_.find(text[0]);
    if (it == singleCharZhengmaCodeHints_.end()) {
        return L"";
    }

    return it->second;
}

bool TextService::IsTextInGB2312Cached(const std::wstring& text) const {
    const auto cacheIt = gb2312TextCache_.find(text);
    if (cacheIt != gb2312TextCache_.end()) {
        return cacheIt->second;
    }

    const bool result = IsTextInGB2312(text);
    if (gb2312TextCache_.size() > 8192) {
        gb2312TextCache_.clear();
    }
    gb2312TextCache_.emplace(text, result);
    return result;
}

size_t TextService::FindPreferredSelectionIndexForPage(size_t pageIndex) const {
    if (allCandidates_.empty()) {
        return 0;
    }

    const size_t start = pageIndex * pageSize_;
    if (start >= allCandidates_.size()) {
        return 0;
    }
    const size_t end = std::min(start + pageSize_, allCandidates_.size());
    if (end <= start) {
        return 0;
    }

    auto scoreCandidate = [this](const CandidateItem& candidate) {
        int score = 0;
        const bool exactFull = !compositionCode_.empty() && candidate.exactMatch;
        if (exactFull) {
            score += 1000;
        }
        if (candidate.boostedUser) {
            score += 300;
        }
        if (candidate.boostedContext) {
            score += 200;
        }
        if (candidate.boostedLearned) {
            score += 100;
        }
        return score;
    };

    size_t bestIndex = 0;
    int bestScore = scoreCandidate(allCandidates_[start]);
    for (size_t i = start + 1; i < end; ++i) {
        const int score = scoreCandidate(allCandidates_[i]);
        if (score > bestScore) {
            bestScore = score;
            bestIndex = i - start;
        }
    }

    return bestIndex;
}

bool TextService::TryFindExactCommitCandidateIndex(size_t& outIndex) const {
    if (compositionCode_.empty()) {
        return false;
    }

    for (size_t i = 0; i < allCandidates_.size(); ++i) {
        const CandidateItem& candidate = allCandidates_[i];
        if (!candidate.exactMatch) {
            continue;
        }

        outIndex = i;
        return true;
    }

    return false;
}

bool TextService::TryFindUniqueExactCommitCandidateIndex(size_t& outIndex) const {
    if (compositionCode_.empty()) {
        return false;
    }

    bool found = false;
    for (size_t i = 0; i < allCandidates_.size(); ++i) {
        const CandidateItem& candidate = allCandidates_[i];
        if (!candidate.exactMatch) {
            continue;
        }

        if (found) {
            return false;
        }

        outIndex = i;
        found = true;
    }

    return found;
}

std::wstring TextService::MakeContextAssociationKey(const std::wstring& prevText, const std::wstring& nextText) {
    return prevText + L"\t" + nextText;
}

bool TextService::LoadContextAssociationFromFile(const std::wstring& filePath) {
    contextAssociationScores_.clear();

    std::ifstream input(filePath, std::ios::in | std::ios::binary);
    if (!input) {
        return false;
    }

    std::string line;
    while (std::getline(input, line)) {
        if (line.empty() || line[0] == '#') {
            continue;
        }

        std::istringstream iss(line);
        std::string prevUtf8;
        std::string nextUtf8;
        std::uint64_t score = 0;
        if (!(iss >> prevUtf8 >> nextUtf8 >> score)) {
            continue;
        }

        const std::wstring prevText = Utf8ToWideText(prevUtf8);
        const std::wstring nextText = Utf8ToWideText(nextUtf8);
        if (prevText.empty() || nextText.empty() || score == 0) {
            continue;
        }

        contextAssociationScores_[MakeContextAssociationKey(prevText, nextText)] = score;
    }

    return true;
}

bool TextService::LoadContextAssociationBlacklistFromFile(const std::wstring& filePath) {
    contextAssociationBlacklist_.clear();

    std::ifstream input(filePath, std::ios::in | std::ios::binary);
    if (!input) {
        return false;
    }

    std::string line;
    while (std::getline(input, line)) {
        if (line.empty() || line[0] == '#') {
            continue;
        }

        std::istringstream iss(line);
        std::string prevUtf8;
        std::string nextUtf8;
        if (!(iss >> prevUtf8 >> nextUtf8)) {
            continue;
        }

        const std::wstring prevText = Utf8ToWideText(prevUtf8);
        const std::wstring nextText = Utf8ToWideText(nextUtf8);
        if (prevText.empty() || nextText.empty()) {
            continue;
        }

        contextAssociationBlacklist_.insert(MakeContextAssociationKey(prevText, nextText));
    }

    return true;
}

bool TextService::SaveContextAssociationToFile(const std::wstring& filePath) const {
    std::ofstream output(filePath, std::ios::out | std::ios::binary | std::ios::trunc);
    if (!output) {
        return false;
    }

    for (const auto& pair : contextAssociationScores_) {
        const size_t split = pair.first.find(L'\t');
        if (split == std::wstring::npos) {
            continue;
        }

        const std::wstring prevText = pair.first.substr(0, split);
        const std::wstring nextText = pair.first.substr(split + 1);
        output << WideToUtf8Text(prevText) << ' ' << WideToUtf8Text(nextText) << ' ' << pair.second << '\n';
    }

    return true;
}

bool TextService::SaveContextAssociationBlacklistToFile(const std::wstring& filePath) const {
    std::ofstream output(filePath, std::ios::out | std::ios::binary | std::ios::trunc);
    if (!output) {
        return false;
    }

    for (const auto& key : contextAssociationBlacklist_) {
        const size_t split = key.find(L'\t');
        if (split == std::wstring::npos) {
            continue;
        }

        const std::wstring prevText = key.substr(0, split);
        const std::wstring nextText = key.substr(split + 1);
        output << WideToUtf8Text(prevText) << ' ' << WideToUtf8Text(nextText) << '\n';
    }

    return true;
}

void TextService::RecordContextAssociation(const std::wstring& prevText, const std::wstring& nextText, std::uint64_t boost) {
    if (!contextAssociationEnabled_ || prevText.empty() || nextText.empty() || boost == 0) {
        return;
    }

    const std::wstring key = MakeContextAssociationKey(prevText, nextText);
    if (contextAssociationBlacklist_.find(key) != contextAssociationBlacklist_.end()) {
        return;
    }
    auto it = contextAssociationScores_.find(key);
    if (it == contextAssociationScores_.end()) {
        contextAssociationScores_[key] = boost;
    } else {
        it->second += boost;
    }

    while (static_cast<int>(contextAssociationScores_.size()) > contextAssociationMaxEntries_) {
        auto minIt = contextAssociationScores_.begin();
        for (auto scan = contextAssociationScores_.begin(); scan != contextAssociationScores_.end(); ++scan) {
            if (scan->second < minIt->second) {
                minIt = scan;
            }
        }
        contextAssociationScores_.erase(minIt);
    }
}

std::uint64_t TextService::QueryContextAssociationScore(const std::wstring& prevText, const std::wstring& nextText) const {
    if (!contextAssociationEnabled_ || prevText.empty() || nextText.empty()) {
        return 0;
    }

    const std::wstring key = MakeContextAssociationKey(prevText, nextText);

    if (contextAssociationBlacklist_.find(key) != contextAssociationBlacklist_.end()) {
        return 0;
    }

    const auto it = contextAssociationScores_.find(key);
    if (it == contextAssociationScores_.end()) {
        return 0;
    }

    return it->second;
}

void TextService::UpdateCandidateWindow() {
    const LONGLONG updateWindowStartCounter = QueryPerfCounterValue();
    if (compositionCode_.empty()) {
        candidateWindow_.Hide();
        return;
    }

    if (!candidateWindow_.EnsureCreated()) {
        Trace(L"UpdateCandidateWindow: candidate window create failed");
        return;
    }

    POINT anchor = {};
    const POINT* anchorPtr = nullptr;
    bool hasFocusedContext = false;
    const ULONGLONG nowTick = GetTickCount64();

    if (candidateWindow_.IsVisible() &&
        hasRecentAnchor_ &&
        nowTick >= lastAnchorTick_ &&
        (nowTick - lastAnchorTick_) <= kAnchorFastReuseWindowMs) {
        anchor = lastAnchor_;
        anchorPtr = &anchor;
        hasFocusedContext = true;
    }

    ITfContext* focusedContext = nullptr;
    if (!hasFocusedContext && TryGetFocusedContext(threadMgr_, &focusedContext)) {
        hasFocusedContext = true;
        if (TryGetCaretScreenPointFromContext(focusedContext, clientId_, anchor)) {
            anchorPtr = &anchor;
            hasRecentAnchor_ = true;
            lastAnchor_ = anchor;
            lastAnchorTick_ = nowTick;
        }
        focusedContext->Release();
    }

    if (!hasFocusedContext) {
        // Defensive guard: if TSF focus is gone, clear stale composition UI immediately.
        ClearComposition();
        candidateWindow_.Hide();
        Trace(L"UpdateCandidateWindow: no focused context, hide candidate window");
        return;
    }

    if (anchorPtr == nullptr) {
        POINT fallbackAnchor = {};
        if (TryGetCaretScreenPoint(fallbackAnchor)) {
            bool useCachedAnchor = false;
            if (hasRecentAnchor_) {
                const ULONGLONG nowTick = GetTickCount64();
                if (nowTick >= lastAnchorTick_ && (nowTick - lastAnchorTick_) <= kAnchorReuseWindowMs) {
                    const int distance = std::abs(fallbackAnchor.x - lastAnchor_.x) + std::abs(fallbackAnchor.y - lastAnchor_.y);
                    useCachedAnchor = distance > 420;
                }
            }

            anchor = useCachedAnchor ? lastAnchor_ : fallbackAnchor;
            anchorPtr = &anchor;
        }
        else if (hasRecentAnchor_) {
            const ULONGLONG nowTick = GetTickCount64();
            if (nowTick >= lastAnchorTick_ && (nowTick - lastAnchorTick_) <= kAnchorReuseWindowMs) {
                anchor = lastAnchor_;
                anchorPtr = &anchor;
            }
        }
    }

    if (anchorPtr != nullptr) {
        hasRecentAnchor_ = true;
        lastAnchor_ = *anchorPtr;
        lastAnchorTick_ = nowTick;
    }

    const auto& pageCandidates = GetCurrentPageCandidates();
    candidateWindow_.Update(
        compositionCode_,
        pageCandidates,
        pageIndex_,
        GetTotalPages(),
        allCandidates_.size(),
        selectedIndexInPage_,
        pageIndex_ * pageSize_ + selectedIndexInPage_,
        chineseMode_,
        fullShapeMode_,
        anchorPtr);
    TraceLatencySample(
        L"candidate-window code=" + compositionCode_ +
            L" visible=" + std::to_wstring(pageCandidates.empty() ? 0 : 1) +
            L" page=" + std::to_wstring(pageIndex_),
        updateWindowStartCounter,
        QueryPerfCounterValue());
}

const std::vector<CandidateWindow::DisplayCandidate>& TextService::GetCurrentPageCandidates() const {
    static const std::vector<CandidateWindow::DisplayCandidate> kEmptyPage;
    if (allCandidates_.empty()) {
        return kEmptyPage;
    }

    const size_t start = pageIndex_ * pageSize_;
    if (start >= allCandidates_.size()) {
        return kEmptyPage;
    }

    if (pageCandidatesCacheRevision_ == candidatesRevision_ &&
        pageCandidatesCachePageIndex_ == pageIndex_ &&
        pageCandidatesCachePageSize_ == pageSize_ &&
        pageCandidatesCacheCode_ == compositionCode_) {
        return pageCandidatesCache_;
    }

    const size_t end = std::min(start + pageSize_, allCandidates_.size());
    pageCandidatesCache_.clear();
    pageCandidatesCache_.reserve(end - start);
    for (size_t i = start; i < end; ++i) {
        CandidateWindow::DisplayCandidate candidate;
        candidate.text = allCandidates_[i].text;
        candidate.code = allCandidates_[i].code;
        if (candidate.code.empty()) {
            candidate.code = allCandidates_[i].commitCode;
        }
        if (candidate.code.empty()) {
            candidate.code = compositionCode_;
        }

        candidate.boostedUser = allCandidates_[i].boostedUser;
        candidate.boostedLearned = allCandidates_[i].boostedLearned;
        candidate.boostedContext = allCandidates_[i].boostedContext;
        candidate.fromAutoPhrase = allCandidates_[i].fromAutoPhrase;
        candidate.fromSessionAutoPhrase = allCandidates_[i].fromSessionAutoPhrase;
        candidate.consumedLength = allCandidates_[i].consumedLength;
        pageCandidatesCache_.push_back(std::move(candidate));
    }

    pageCandidatesCacheRevision_ = candidatesRevision_;
    pageCandidatesCachePageIndex_ = pageIndex_;
    pageCandidatesCachePageSize_ = pageSize_;
    pageCandidatesCacheCode_ = compositionCode_;
    return pageCandidatesCache_;
}

size_t TextService::GetCurrentPageCandidateCount() const {
    if (allCandidates_.empty()) {
        return 0;
    }

    const size_t start = pageIndex_ * pageSize_;
    if (start >= allCandidates_.size()) {
        return 0;
    }

    return std::min(pageSize_, allCandidates_.size() - start);
}

void TextService::InvalidatePageCandidatesCache() {
    ++candidatesRevision_;
}

size_t TextService::GetTotalPages() const {
    if (allCandidates_.empty()) {
        return 0;
    }

    return (allCandidates_.size() + pageSize_ - 1) / pageSize_;
}

void TextService::ScheduleDeferredCandidateExpansion() {
    if (compositionCode_.empty() ||
        candidatesFullyExpanded_ ||
        !candidateWindow_.IsVisible() ||
        compositionCode_.size() <= 2) {
        deferredExpansionCode_.clear();
        deferredExpansionDueTick_ = 0;
        candidateWindow_.CancelAsyncPoll();
        return;
    }

    deferredExpansionCode_ = compositionCode_;
    deferredExpansionDueTick_ = GetTickCount64() + 45ULL;
    candidateWindow_.ScheduleAsyncPoll(45);
}

void TextService::RunDeferredCandidateExpansion() {
    if (deferredExpansionCode_.empty()) {
        return;
    }

    if (compositionCode_ != deferredExpansionCode_ || compositionCode_.empty() || candidatesFullyExpanded_) {
        deferredExpansionCode_.clear();
        deferredExpansionDueTick_ = 0;
        return;
    }

    const ULONGLONG now = GetTickCount64();
    if (now < deferredExpansionDueTick_) {
        candidateWindow_.ScheduleAsyncPoll(static_cast<UINT>(std::max<ULONGLONG>(1, deferredExpansionDueTick_ - now)));
        return;
    }

    deferredExpansionCode_.clear();
    deferredExpansionDueTick_ = 0;
    RefreshCandidates(true);
}

bool TextService::TryRestoreCachedCandidatesForCode(const std::wstring& code) {
    const LONGLONG restoreStartCounter = QueryPerfCounterValue();
    if (code.empty()) {
        return false;
    }

    const auto it = compositionCandidateCache_.find(code);
    if (it == compositionCandidateCache_.end()) {
        return false;
    }

    compositionCode_ = code;
    allCandidates_ = it->second.candidates;
    candidatesFullyExpanded_ = it->second.fullyExpanded;
    InvalidatePageCandidatesCache();
    pageIndex_ = 0;
    selectedIndexInPage_ = FindPreferredSelectionIndexForPage(pageIndex_);
    emptyCandidateAlerted_ = allCandidates_.empty();
    UpdateCandidateWindow();
    if (!candidatesFullyExpanded_) {
        ScheduleDeferredCandidateExpansion();
    }
    TraceLatencySample(
        L"restore-cache code=" + code + L" count=" + std::to_wstring(allCandidates_.size()),
        restoreStartCounter,
        QueryPerfCounterValue());
    return true;
}

void TextService::CacheCurrentCandidatesForCode() {
    if (compositionCode_.empty()) {
        return;
    }

    CachedCandidateState state;
    state.candidates = allCandidates_;
    state.fullyExpanded = candidatesFullyExpanded_;

    const bool alreadyPresent = compositionCandidateCache_.find(compositionCode_) != compositionCandidateCache_.end();
    compositionCandidateCache_[compositionCode_] = std::move(state);
    if (!alreadyPresent) {
        compositionCandidateCacheOrder_.push_back(compositionCode_);
    }

    while (compositionCandidateCacheOrder_.size() > 12) {
        const std::wstring oldestCode = compositionCandidateCacheOrder_.front();
        compositionCandidateCacheOrder_.pop_front();
        compositionCandidateCache_.erase(oldestCode);
    }
}

void TextService::ClearComposition() {
    compositionCode_.clear();
    allCandidates_.clear();
    candidatesFullyExpanded_ = false;
    compositionCandidateCache_.clear();
    compositionCandidateCacheOrder_.clear();
    deferredExpansionCode_.clear();
    deferredExpansionDueTick_ = 0;
    candidateWindow_.CancelAsyncPoll();
    InvalidatePageCandidatesCache();
    pageIndex_ = 0;
    selectedIndexInPage_ = 0;
    emptyCandidateAlerted_ = false;
    pageBoundaryDirection_ = 0;
    pageBoundaryHitCount_ = 0;
}

void TextService::LearnPhraseFromRecentCommits(const std::wstring& committedCode, const std::wstring& committedText) {
    if (committedCode.empty() || committedText.empty()) {
        return;
    }

    const ULONGLONG now = GetTickCount64();

    while (!recentCommits_.empty() && (now - recentCommits_.front().tick) > 60000ULL) {
        recentCommits_.pop_front();
    }

    if (!recentCommits_.empty()) {
        const CommitHistoryItem& prev = recentCommits_.back();
        const bool closeEnough = (now - prev.tick) <= 8000ULL;
        if (closeEnough) {
            RecordContextAssociation(prev.text, committedText, 1);
            if (!contextAssocPath_.empty()) {
                const std::string snapshot = BuildContextAssociationFileContent();
                {
                    std::lock_guard<std::mutex> lock(userDataWriteMutex_);
                    pendingContextAssocWrite_.path = contextAssocPath_;
                    pendingContextAssocWrite_.content = snapshot;
                    pendingContextAssocWrite_.deleteIfEmpty = false;
                    pendingContextAssocWrite_.generation = ++nextUserDataWriteGeneration_;
                }
                userDataStampRefreshPending_ = true;
                userDataWriteCv_.notify_all();
            }
        }
    }

    UpdateSessionAutoPhraseHistory(committedText, now);

    CommitHistoryItem item;
    item.code = committedCode;
    item.text = committedText;
    item.tick = now;
    recentCommits_.push_back(std::move(item));
    while (recentCommits_.size() > 24) {
        recentCommits_.pop_front();
    }
}

bool TextService::CommitCandidateByGlobalIndex(ITfContext* context, size_t globalIndex, std::uint64_t freqBoost) {
    if (globalIndex >= allCandidates_.size()) {
        Trace(L"select: out of range");
        return false;
    }

    const CandidateItem& candidate = allCandidates_[globalIndex];
    const std::wstring& textToCommit = candidate.text;
    const std::wstring freqCode = candidate.commitCode.empty() ? compositionCode_ : candidate.commitCode;
    const size_t consumedLength = candidate.consumedLength == 0 ? compositionCode_.size() : std::min(candidate.consumedLength, compositionCode_.size());
    const bool keepComposing = consumedLength < compositionCode_.size();
    const std::wstring remainingCode = keepComposing ? compositionCode_.substr(consumedLength) : L"";

    Trace(L"select: candidate=" + textToCommit);
    if (!CommitText(context, textToCommit)) {
        Trace(L"commit(index) failed text=" + textToCommit);
        return false;
    }

    const ULONGLONG nowTick = GetTickCount64();
    const std::wstring selectedKey = freqCode + L"\t" + textToCommit;
    if (candidate.fromSessionAutoPhrase) {
        PromoteSessionAutoPhrase(textToCommit);
    } else if (candidate.fromAutoPhrase && !freqCode.empty()) {
        engine_.AddUserEntry(freqCode, textToCommit);
        AppendPhraseReviewEntry(freqCode, textToCommit, L"auto-commit");
        MarkAutoPhraseDictionaryDirty();
    }

    if (candidate.fromAutoPhrase || candidate.fromSessionAutoPhrase) {
        const bool withinWindow =
            autoPhraseSelectedTick_ != 0 &&
            nowTick >= autoPhraseSelectedTick_ &&
            (nowTick - autoPhraseSelectedTick_) <= 20000ULL;
        if (withinWindow && selectedKey == lastAutoPhraseSelectedKey_) {
            ++autoPhraseSelectedStreak_;
        } else {
            lastAutoPhraseSelectedKey_ = selectedKey;
            autoPhraseSelectedStreak_ = 1;
        }
        autoPhraseSelectedTick_ = nowTick;
    } else {
        lastAutoPhraseSelectedKey_.clear();
        autoPhraseSelectedStreak_ = 0;
        autoPhraseSelectedTick_ = 0;
    }

    engine_.RecordCommit(freqCode, textToCommit, freqBoost);
    LearnPhraseFromRecentCommits(freqCode, textToCommit);
    MarkFrequencyDataDirty();
    FlushPendingUserDataIfNeeded(false);

    if (keepComposing) {
        compositionCode_ = remainingCode;
        RefreshCandidates();
        Trace(L"commit(index)=" + textToCommit + L" continue=" + remainingCode);
    } else {
        ClearComposition();
        candidateWindow_.Hide();
        Trace(L"commit(index)=" + textToCommit);
    }

    return true;
}

bool TextService::PinCandidateByGlobalIndex(size_t globalIndex) {
    if (globalIndex >= allCandidates_.size()) {
        Trace(L"pin: out of range");
        return false;
    }

    const CandidateItem& candidate = allCandidates_[globalIndex];
    const std::wstring pinCode = !compositionCode_.empty()
        ? compositionCode_
        : (candidate.commitCode.empty() ? candidate.code : candidate.commitCode);
    if (pinCode.empty()) {
        return false;
    }

    engine_.PinEntry(pinCode, candidate.text);
    MarkAutoPhraseDictionaryDirty();
    FlushPendingUserDataIfNeeded(false);

    RefreshCandidates();
    Trace(L"pin code=" + pinCode + L" text=" + candidate.text);
    return true;
}

bool TextService::BlockCandidateByGlobalIndex(size_t globalIndex) {
    if (globalIndex >= allCandidates_.size()) {
        Trace(L"block: out of range");
        return false;
    }

    const CandidateItem& candidate = allCandidates_[globalIndex];
    const std::wstring blockCode = !compositionCode_.empty()
        ? compositionCode_
        : (candidate.commitCode.empty() ? candidate.code : candidate.commitCode);
    if (blockCode.empty()) {
        return false;
    }

    engine_.BlockEntry(blockCode, candidate.text);
    if (!blockedEntriesPath_.empty()) {
        const std::string snapshot = engine_.BuildBlockedEntriesFileContent();
        {
            std::lock_guard<std::mutex> lock(userDataWriteMutex_);
            pendingBlockedEntriesWrite_.path = blockedEntriesPath_;
            pendingBlockedEntriesWrite_.content = snapshot;
            pendingBlockedEntriesWrite_.deleteIfEmpty = false;
            pendingBlockedEntriesWrite_.append = false;
            pendingBlockedEntriesWrite_.generation = ++nextUserDataWriteGeneration_;
        }
        userDataStampRefreshPending_ = true;
        userDataWriteCv_.notify_all();
    }
    MarkAutoPhraseDictionaryDirty();
    MarkFrequencyDataDirty();
    FlushPendingUserDataIfNeeded(false);

    RefreshCandidates();
    Trace(L"block code=" + blockCode + L" text=" + candidate.text);
    return true;
}

bool TextService::ShowStatusMenu() {
    HMENU menu = CreatePopupMenu();
    if (menu == nullptr) {
        return false;
    }

    AppendMenuW(menu, MF_STRING | (chineseMode_ ? MF_CHECKED : 0), kMenuToggleChinese, L"Chinese Mode (Ctrl+Space)");
    AppendMenuW(menu, MF_STRING | (fullShapeMode_ ? MF_CHECKED : 0), kMenuToggleShape, L"Full Shape (F10)");
    AppendMenuW(menu, MF_STRING, kMenuOpenConfig, L"Open Config...");
    AppendMenuW(menu, MF_STRING, kMenuOpenSystemSettings, L"Open System Input Settings...");
    AppendMenuW(menu, MF_STRING, kMenuOpenRuntimeLog, L"Open Runtime Log...");
    AppendMenuW(menu, MF_SEPARATOR, 0, nullptr);
    AppendMenuW(menu, MF_STRING | MF_GRAYED, kMenuHelp, L"Config: Ctrl+F2 (or Ctrl+Shift+F2) | Mode: Ctrl+Shift+F3/F4 | Manual: Ctrl+Shift+M | Page: [ ] , . PgUp PgDn | Pin: Ctrl+1-9 | Delete: Ctrl+Del");

    POINT pt = {120, 120};
    TryGetCaretScreenPoint(pt);

    HWND menuOwner = GetForegroundWindow();
    if (menuOwner == nullptr) {
        menuOwner = GetDesktopWindow();
    }
    SetForegroundWindow(menuOwner);

    const UINT cmd = TrackPopupMenu(menu, TPM_RETURNCMD | TPM_NONOTIFY | TPM_RIGHTBUTTON, pt.x, pt.y + 12, 0, menuOwner, nullptr);
    DestroyMenu(menu);

    if (cmd == 0) {
        Trace(L"status menu cmd=0");
        return false;
    }

    if (cmd == kMenuToggleChinese) {
        chineseMode_ = !chineseMode_;
        if (!chineseMode_) {
            ClearComposition();
            candidateWindow_.Hide();
        } else {
            UpdateCandidateWindow();
        }
        Trace(std::wstring(L"menu mode=") + (chineseMode_ ? L"ZH" : L"EN"));
        return true;
    }

    if (cmd == kMenuToggleShape) {
        fullShapeMode_ = !fullShapeMode_;
        if (!compositionCode_.empty()) {
            UpdateCandidateWindow();
        }
        Trace(std::wstring(L"menu shape=") + (fullShapeMode_ ? L"FULL" : L"HALF"));
        return true;
    }

    if (cmd == kMenuOpenConfig) {
        const bool launched = LaunchConfigExecutable();
        Trace(std::wstring(L"menu open config=") + (launched ? L"1" : L"0"));
        return true;
    }

    if (cmd == kMenuOpenSystemSettings) {
        const bool launched = LaunchSystemImeSettings();
        Trace(std::wstring(L"menu open system settings=") + (launched ? L"1" : L"0"));
        return true;
    }

    if (cmd == kMenuOpenRuntimeLog) {
        const bool launched = LaunchRuntimeLogFile();
        Trace(std::wstring(L"menu open runtime log=") + (launched ? L"1" : L"0"));
        return true;
    }

    return true;
}

bool TextService::CommitAsciiKey(ITfContext* context, WPARAM wParam, LPARAM lParam) {
    wchar_t c = 0;
    if (!TryBuildDirectAsciiCommitChar(wParam, lParam, c)) {
        return false;
    }

    if (fullShapeMode_) {
        if (c == L' ') {
            c = static_cast<wchar_t>(0x3000);
        } else if (c >= 0x21 && c <= 0x7E) {
            c = static_cast<wchar_t>(c + 0xFEE0);
        }
    }

    const bool committed = CommitText(context, std::wstring(1, c));
    if (committed) {
        RecordSessionAutoPhraseBreak();
    }
    return committed;
}

bool TextService::PromoteSelectedCandidateToManualEntry() {
    if (compositionCode_.empty()) {
        return false;
    }

    const size_t pageCandidateCount = GetCurrentPageCandidateCount();
    if (pageCandidateCount == 0) {
        return false;
    }

    const size_t safeIndexInPage = std::min(selectedIndexInPage_, pageCandidateCount - 1);
    const size_t globalIndex = pageIndex_ * pageSize_ + safeIndexInPage;
    if (globalIndex >= allCandidates_.size()) {
        return false;
    }

    const CandidateItem& candidate = allCandidates_[globalIndex];
    const std::wstring manualCode = candidate.commitCode.empty() ? compositionCode_ : candidate.commitCode;
    if (manualCode.empty()) {
        return false;
    }

    const auto normalizePhraseCode = [](std::wstring code) {
        std::transform(
            code.begin(),
            code.end(),
            code.begin(),
            [](wchar_t ch) {
                return static_cast<wchar_t>(std::towlower(static_cast<wint_t>(ch)));
            });
        code.erase(
            std::remove_if(
                code.begin(),
                code.end(),
                [](wchar_t ch) {
                    return ch < L'a' || ch > L'z';
                }),
            code.end());
        return code;
    };

    std::vector<std::wstring> manualCodes;
    manualCodes.push_back(manualCode);
    if (candidate.text.size() >= 2) {
        std::vector<std::wstring> phraseCodes;
        if (phraseBuildEngine_.TryBuildPhraseCodes(candidate.text, phraseCodes)) {
            for (std::wstring& phraseCode : phraseCodes) {
                phraseCode = normalizePhraseCode(std::move(phraseCode));
                if (phraseCode.empty()) {
                    continue;
                }
                if (std::find(manualCodes.begin(), manualCodes.end(), phraseCode) == manualCodes.end()) {
                    manualCodes.push_back(std::move(phraseCode));
                }
            }
        }
    }

    bool changed = false;
    for (const std::wstring& code : manualCodes) {
        changed = engine_.PinEntry(code, candidate.text) || changed;
        AppendPhraseReviewEntry(code, candidate.text, L"manual");
    }
    if (changed || !manualCodes.empty()) {
        MarkAutoPhraseDictionaryDirty();
        FlushPendingUserDataIfNeeded(false);
    }
    RefreshCandidates();
    Trace(L"manual phrase code=" + manualCode + L" text=" + candidate.text + L" variants=" + std::to_wstring(manualCodes.size()));
    return changed || !manualCodes.empty();
}

void TextService::AppendPhraseReviewEntry(const std::wstring& code, const std::wstring& text, const wchar_t* sourceTag) {
    if (manualPhraseReviewPath_.empty() || code.empty() || text.empty() || sourceTag == nullptr) {
        return;
    }

    SYSTEMTIME now = {};
    GetLocalTime(&now);
    wchar_t timestamp[40] = {};
    swprintf_s(timestamp, L"%04u-%02u-%02uT%02u:%02u:%02u", now.wYear, now.wMonth, now.wDay, now.wHour, now.wMinute, now.wSecond);

    const std::string line =
        WideToUtf8Text(code) + " " +
        WideToUtf8Text(text) + " " +
        WideToUtf8Text(sourceTag) + " " +
        WideToUtf8Text(timestamp) + "\n";
    QueueManualPhraseReviewAppend(line);
}

bool TextService::CommitText(ITfContext* context, const std::wstring& text) {
    (void)context;
    if (text.empty()) {
        return false;
    }

    const bool sent = SendUnicodeTextWithInput(text);
    if (sent) {
        bool hasHan = false;
        bool hasNonHan = false;
        for (wchar_t ch : text) {
            if (IsHanCharacter(ch)) {
                hasHan = true;
            } else {
                hasNonHan = true;
            }
        }

        if (!hasHan || hasNonHan) {
            RecordSessionAutoPhraseBreak();
        }
    }
    Trace(std::wstring(L"CommitText direct SendInput=") + (sent ? L"1" : L"0") + L" text=" + text);
    return sent;
}

bool TextService::IsAlphaKey(WPARAM wParam) {
    return (wParam >= L'A' && wParam <= L'Z') || (wParam >= L'a' && wParam <= L'z');
}

wchar_t TextService::ToLowerAlpha(WPARAM wParam) {
    wchar_t c = static_cast<wchar_t>(wParam);
    if (c >= L'A' && c <= L'Z') {
        c = static_cast<wchar_t>(c - L'A' + L'a');
    }
    return c;
}

void TextService::LoadSettings() {
    std::filesystem::path settingsPath;

    wchar_t localAppData[MAX_PATH] = {};
    const DWORD localLen = GetEnvironmentVariableW(L"LOCALAPPDATA", localAppData, MAX_PATH);
    if (localLen > 0 && localLen < MAX_PATH) {
        const auto localPath = std::filesystem::path(std::wstring(localAppData, localLen)) / L"yuninput" / L"settings.json";
        if (std::filesystem::exists(localPath)) {
            settingsPath = localPath;
        }
    }

    if (settingsPath.empty() && !userDataDir_.empty()) {
        const auto roamingPath = std::filesystem::path(userDataDir_) / L"settings.json";
        if (std::filesystem::exists(roamingPath)) {
            settingsPath = roamingPath;
        }
    }

    if (settingsPath.empty()) {
        return;
    }

    const std::wstring text = ReadTextFile(settingsPath);
    if (text.empty()) {
        return;
    }

    bool boolValue = false;
    if (ExtractBool(text, L"chinese_mode", boolValue)) {
        chineseMode_ = boolValue;
    }
    if (ExtractBool(text, L"full_shape", boolValue)) {
        fullShapeMode_ = boolValue;
    }
    if (ExtractBool(text, L"chinese_punctuation", boolValue)) {
        chinesePunctuation_ = boolValue;
    }
    if (ExtractBool(text, L"smart_symbol_pairs", boolValue)) {
        smartSymbolPairs_ = boolValue;
    }
    if (ExtractBool(text, L"auto_commit_unique_exact", boolValue)) {
        autoCommitUniqueExact_ = boolValue;
    }
    if (ExtractBool(text, L"empty_candidate_beep", boolValue)) {
        emptyCandidateBeep_ = boolValue;
    }
    if (ExtractBool(text, L"tab_navigation", boolValue)) {
        tabNavigation_ = boolValue;
    }
    if (ExtractBool(text, L"enter_exact_priority", boolValue)) {
        enterExactPriority_ = boolValue;
    }
    if (ExtractBool(text, L"context_association_enabled", boolValue)) {
        contextAssociationEnabled_ = boolValue;
    }

    int pageSizeValue = static_cast<int>(kDefaultPageSize);
    if (ExtractInt(text, L"candidate_page_size", pageSizeValue)) {
        if (pageSizeValue < 1) {
            pageSizeValue = 1;
        }
        if (pageSizeValue > 6) {
            pageSizeValue = 6;
        }
        pageSize_ = static_cast<size_t>(pageSizeValue);
    }

    int autoCommitLength = autoCommitMinCodeLength_;
    if (ExtractInt(text, L"auto_commit_min_code_length", autoCommitLength)) {
        if (autoCommitLength < 2) {
            autoCommitLength = 2;
        }
        if (autoCommitLength > 8) {
            autoCommitLength = 8;
        }
        autoCommitMinCodeLength_ = autoCommitLength;
    }

    int contextAssocMax = contextAssociationMaxEntries_;
    if (ExtractInt(text, L"context_association_max_entries", contextAssocMax)) {
        if (contextAssocMax < 1000) {
            contextAssocMax = 1000;
        }
        if (contextAssocMax > 50000) {
            contextAssocMax = 50000;
        }
        contextAssociationMaxEntries_ = contextAssocMax;
    }

    std::wstring hotkeyValue;
    if (ExtractString(text, L"toggle_hotkey", hotkeyValue)) {
        if (_wcsicmp(hotkeyValue.c_str(), L"F8") == 0) {
            toggleHotkey_ = ToggleHotkey::F8;
        } else if (_wcsicmp(hotkeyValue.c_str(), L"Ctrl+Space") == 0) {
            toggleHotkey_ = ToggleHotkey::CtrlSpace;
        } else {
            toggleHotkey_ = ToggleHotkey::F9;
        }
    }

    std::wstring dictionaryProfileValue;
    if (ExtractString(text, L"dictionary_profile", dictionaryProfileValue)) {
        if (_wcsicmp(dictionaryProfileValue.c_str(), L"zhengma-large-pinyin") == 0 ||
            _wcsicmp(dictionaryProfileValue.c_str(), L"pinyin") == 0) {
            dictionaryProfile_ = DictionaryProfile::ZhengmaLargePinyin;
        } else if (_wcsicmp(dictionaryProfileValue.c_str(), L"zhengma-all") == 0 ||
                   _wcsicmp(dictionaryProfileValue.c_str(), L"zhengma-large") == 0) {
            dictionaryProfile_ = DictionaryProfile::ZhengmaLarge;
        } else {
            dictionaryProfile_ = DictionaryProfile::ZhengmaLarge;
        }
    }
}

bool TextService::PersistDictionaryProfileSetting() const {
    std::filesystem::path settingsPath;

    wchar_t localAppData[MAX_PATH] = {};
    const DWORD localLen = GetEnvironmentVariableW(L"LOCALAPPDATA", localAppData, MAX_PATH);
    if (localLen > 0 && localLen < MAX_PATH) {
        settingsPath = std::filesystem::path(std::wstring(localAppData, localLen)) / L"yuninput" / L"settings.json";
    } else if (!userDataDir_.empty()) {
        settingsPath = std::filesystem::path(userDataDir_) / L"settings.json";
    } else {
        return false;
    }

    std::error_code ec;
    std::filesystem::create_directories(settingsPath.parent_path(), ec);

    const std::wstring profileValue = GetDictionaryProfileName(dictionaryProfile_);
    std::wstring text = ReadTextFile(settingsPath);

    if (text.empty()) {
        text = L"{\n  \"dictionary_profile\": \"" + profileValue + L"\"\n}\n";
    } else {
        const std::wregex profileRe(L"\"dictionary_profile\"\\s*:\\s*\"[^\"]*\"");
        if (std::regex_search(text, profileRe)) {
            text = std::regex_replace(text, profileRe, L"\"dictionary_profile\": \"" + profileValue + L"\"");
        } else {
            const size_t closePos = text.rfind(L'}');
            if (closePos == std::wstring::npos) {
                text = L"{\n  \"dictionary_profile\": \"" + profileValue + L"\"\n}\n";
            } else {
                size_t insertPos = closePos;
                while (insertPos > 0 && std::iswspace(static_cast<wint_t>(text[insertPos - 1])) != 0) {
                    --insertPos;
                }

                const bool hasAnyField = insertPos > 0 && text[insertPos - 1] != L'{';
                std::wstring insertion;
                if (hasAnyField) {
                    insertion = L",\n  \"dictionary_profile\": \"" + profileValue + L"\"\n";
                } else {
                    insertion = L"\n  \"dictionary_profile\": \"" + profileValue + L"\"\n";
                }
                text.insert(insertPos, insertion);
            }
        }
    }

    std::ofstream output(settingsPath, std::ios::out | std::ios::binary | std::ios::trunc);
    if (!output) {
        return false;
    }

    const std::string utf8 = WideToUtf8Text(text);
    output.write(utf8.data(), static_cast<std::streamsize>(utf8.size()));
    return output.good();
}

void TextService::NotifyDictionaryProfileSwitch(DictionaryProfile profile, bool success) const {
    if (!success) {
        MessageBeep(MB_ICONHAND);
        return;
    }

    int beepCount = 1;
    switch (profile) {
    case DictionaryProfile::ZhengmaLarge:
        beepCount = 1;
        break;
    case DictionaryProfile::ZhengmaLargePinyin:
        beepCount = 2;
        break;
    default:
        beepCount = 1;
        break;
    }

    for (int i = 0; i < beepCount; ++i) {
        MessageBeep(MB_ICONASTERISK);
        if (i + 1 < beepCount) {
            Sleep(70);
        }
    }
}

bool TextService::IsToggleHotkeyPressed(WPARAM wParam) const {
    switch (toggleHotkey_) {
    case ToggleHotkey::F8:
        return wParam == VK_F8;
    case ToggleHotkey::CtrlSpace:
        return wParam == VK_SPACE && (GetKeyState(VK_CONTROL) & 0x8000) != 0;
    case ToggleHotkey::F9:
    default:
        return wParam == VK_F9;
    }
}

STDMETHODIMP TextService::OnSetFocus(BOOL foreground) {
    if (!foreground) {
        const bool hadComposition = !compositionCode_.empty();
        if (hadComposition) {
            ClearComposition();
        }
        const bool hidWindow = candidateWindow_.Hide();
        if (hadComposition || hidWindow) {
            Trace(L"OnSetFocus(BOOL): hide candidate window");
        }
    }
    return S_OK;
}

STDMETHODIMP TextService::OnInitDocumentMgr(ITfDocumentMgr*) {
    return S_OK;
}

STDMETHODIMP TextService::OnUninitDocumentMgr(ITfDocumentMgr* documentMgr) {
    ITfDocumentMgr* focused = nullptr;
    const HRESULT focusHr = (threadMgr_ != nullptr) ? threadMgr_->GetFocus(&focused) : E_FAIL;
    const bool losingFocusedDoc = SUCCEEDED(focusHr) && focused == documentMgr;
    if (focused != nullptr) {
        focused->Release();
    }

    if (losingFocusedDoc) {
        const bool hadComposition = !compositionCode_.empty();
        if (hadComposition) {
            ClearComposition();
        }
        const bool hidWindow = candidateWindow_.Hide();
        if (hadComposition || hidWindow) {
            Trace(L"OnUninitDocumentMgr: hide candidate window");
        }
    }
    return S_OK;
}

STDMETHODIMP TextService::OnSetFocus(ITfDocumentMgr* documentMgrFocus, ITfDocumentMgr* documentMgrPrevFocus) {
    if (documentMgrFocus != documentMgrPrevFocus) {
        const bool hadComposition = !compositionCode_.empty();
        if (hadComposition) {
            ClearComposition();
        }
        const bool hidWindow = candidateWindow_.Hide();
        if (hadComposition || hidWindow) {
            Trace(L"OnSetFocus(DocumentMgr): hide candidate window");
        }

        if (documentMgrFocus != nullptr && chineseMode_) {
            EnsureRuntimeReady();
            ReloadUserDataIfChanged(false);
        }
    }
    return S_OK;
}

STDMETHODIMP TextService::OnPushContext(ITfContext*) {
    return S_OK;
}

STDMETHODIMP TextService::OnPopContext(ITfContext*) {
    return S_OK;
}

STDMETHODIMP TextService::OnTestKeyDown(ITfContext*, WPARAM wParam, LPARAM lParam, BOOL* eaten) {
    if (eaten == nullptr) {
        return E_INVALIDARG;
    }

    const WPARAM key = NormalizeVirtualKey(wParam, lParam);
    const bool verboseKeyTrace = IsVerboseKeyTraceEnabled();
    const bool ctrlPressed =
        (GetKeyState(VK_CONTROL) & 0x8000) != 0 ||
        (GetKeyState(VK_LCONTROL) & 0x8000) != 0 ||
        (GetKeyState(VK_RCONTROL) & 0x8000) != 0 ||
        (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0 ||
        (GetAsyncKeyState(VK_LCONTROL) & 0x8000) != 0 ||
        (GetAsyncKeyState(VK_RCONTROL) & 0x8000) != 0;
    const bool altPressed =
        (GetKeyState(VK_MENU) & 0x8000) != 0 ||
        (GetKeyState(VK_LMENU) & 0x8000) != 0 ||
        (GetKeyState(VK_RMENU) & 0x8000) != 0;
    const bool winPressed =
        (GetKeyState(VK_LWIN) & 0x8000) != 0 ||
        (GetKeyState(VK_RWIN) & 0x8000) != 0;
    const bool shiftPressed = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
    const bool modeSwitchHotkey = ctrlPressed && shiftPressed && (key == VK_F3 || key == VK_F4);
    const bool manualEntryHotkey = ctrlPressed && shiftPressed && (key == L'M' || key == L'm');
    size_t selectionOffset = 0;
    const bool ctrlDigitHotkey = ctrlPressed && TryMapSelectionInputToIndex(key, lParam, selectionOffset);
    const bool ctrlDeleteHotkey = ctrlPressed && key == VK_DELETE;
    const bool ctrlConfigHotkey = ctrlPressed && key == VK_F2;
    const bool shouldBypassGenericCtrlShortcut = ctrlPressed &&
        !IsToggleHotkeyPressed(key) &&
        !modeSwitchHotkey &&
        !manualEntryHotkey &&
        !ctrlConfigHotkey &&
        !ctrlDigitHotkey &&
        !ctrlDeleteHotkey;

    if (shouldBypassGenericCtrlShortcut) {
        *eaten = FALSE;
        return S_OK;
    }

    if (leftShiftTogglePending_ && !IsLeftShiftToggleKey(key, lParam)) {
        leftShiftTogglePending_ = false;
        leftShiftToggleDownTick_ = 0;
    }

    if (IsLeftShiftToggleKey(key, lParam) && !ctrlPressed && !altPressed && !winPressed) {
        const bool wasDown = ((static_cast<UINT>(lParam) >> 30) & 0x1U) != 0;
        if (!wasDown) {
            leftShiftTogglePending_ = true;
            leftShiftToggleDownTick_ = GetTickCount64();
        }
        if (verboseKeyTrace) {
            wchar_t shiftTestLog[220] = {};
            swprintf_s(
                shiftTestLog,
                L"OnTestKeyDown shift raw=0x%04X key=0x%04X pending=%d wasDown=%d chinese=%d fullShape=%d",
                static_cast<unsigned int>(wParam),
                static_cast<unsigned int>(key),
                leftShiftTogglePending_ ? 1 : 0,
                wasDown ? 1 : 0,
                chineseMode_ ? 1 : 0,
                fullShapeMode_ ? 1 : 0);
            Trace(shiftTestLog);
        }
    }

    if (altPressed || winPressed) {
        *eaten = FALSE;
        return S_OK;
    }

    if (!chineseMode_) {
        if (!fullShapeMode_) {
            *eaten = (key == VK_F8 ||
                      key == VK_F9 ||
                      key == VK_F10 ||
                      modeSwitchHotkey ||
                      manualEntryHotkey ||
                      key == VK_F2 ||
                      key == VK_APPS)
                         ? TRUE
                         : FALSE;
            if (verboseKeyTrace &&
                (key == VK_BACK || key == VK_DELETE || key == VK_OEM_MINUS || key == VK_SUBTRACT || key == VK_OEM_PLUS || key == VK_ADD || IsLeftShiftToggleKey(key, lParam))) {
                wchar_t englishBypassLog[240] = {};
                swprintf_s(
                    englishBypassLog,
                    L"OnTestKeyDown EN-half raw=0x%04X key=0x%04X eaten=%d pending=%d",
                    static_cast<unsigned int>(wParam),
                    static_cast<unsigned int>(key),
                    *eaten ? 1 : 0,
                    leftShiftTogglePending_ ? 1 : 0);
                Trace(englishBypassLog);
            }
            return S_OK;
        }

        const bool englishDirectCommitKey =
            IsAlphaKey(key) ||
            (key >= L'0' && key <= L'9') ||
            (key >= VK_NUMPAD0 && key <= VK_NUMPAD9) ||
            key == VK_SPACE ||
            IsPunctuationVirtualKey(key);

        *eaten = (englishDirectCommitKey ||
                  key == VK_F8 ||
                  key == VK_F9 ||
                  key == VK_F10 ||
                  modeSwitchHotkey ||
                  manualEntryHotkey ||
                  key == VK_F2 ||
                  key == VK_APPS)
                     ? TRUE
                     : FALSE;
        return S_OK;
    }

    const bool hasComposition = !compositionCode_.empty();

    const bool minusPageUpKey = IsMinusPageUpKey(key, wParam, lParam);
    const bool plusPageDownKey = IsPlusPageDownKey(key, wParam, lParam);

    *eaten = (IsAlphaKey(key) ||
              (key >= L'0' && key <= L'9') ||
              (key >= L'1' && key <= L'9') ||
              (key >= VK_NUMPAD0 && key <= VK_NUMPAD9) ||
              (tabNavigation_ && key == VK_TAB) ||
              (key == VK_BACK && hasComposition) ||
              (key == VK_DELETE && hasComposition) ||
              (key == VK_SPACE && hasComposition) ||
              (key == VK_RETURN && hasComposition) ||
              key == VK_ESCAPE ||
              key == VK_UP ||
              key == VK_DOWN ||
              key == VK_PRIOR ||
              key == VK_NEXT ||
              key == VK_F8 ||
              key == VK_F9 ||
              key == VK_F10 ||
              modeSwitchHotkey ||
              key == VK_F2 ||
              key == VK_APPS ||
                  key == VK_OEM_1 ||
                  key == VK_OEM_2 ||
                  key == VK_OEM_3 ||
              minusPageUpKey ||
              plusPageDownKey ||
              key == VK_OEM_COMMA ||
              key == VK_OEM_PERIOD ||
              key == VK_OEM_4 ||
                  key == VK_OEM_5 ||
                  key == VK_OEM_7 ||
              key == VK_OEM_6)
                 ? TRUE
                 : FALSE;
    if (verboseKeyTrace &&
        (IsAlphaKey(key) ||
         (key >= L'0' && key <= L'9') ||
         (key >= VK_NUMPAD0 && key <= VK_NUMPAD9) ||
         key == VK_SPACE ||
         key == VK_RETURN ||
         key == VK_F9 ||
         wParam == VK_PROCESSKEY)) {
        wchar_t keyLog[200] = {};
        swprintf_s(keyLog, L"OnTestKeyDown raw=0x%04X key=0x%04X eaten=%d chinese=%d", static_cast<unsigned int>(wParam), static_cast<unsigned int>(key), *eaten ? 1 : 0, chineseMode_ ? 1 : 0);
        Trace(keyLog);
    }
    return S_OK;
}

STDMETHODIMP TextService::OnTestKeyUp(ITfContext*, WPARAM wParam, LPARAM lParam, BOOL* eaten) {
    if (eaten == nullptr) {
        return E_INVALIDARG;
    }

    const WPARAM key = NormalizeVirtualKey(wParam, lParam);
    if (IsLeftShiftToggleKey(key, lParam)) {
        *eaten = TRUE;
        return S_OK;
    }

    *eaten = FALSE;
    return S_OK;
}

STDMETHODIMP TextService::OnKeyDown(ITfContext* context, WPARAM wParam, LPARAM lParam, BOOL* eaten) {
    if (eaten == nullptr) {
        return E_INVALIDARG;
    }

    const WPARAM key = NormalizeVirtualKey(wParam, lParam);
    const bool verboseKeyTrace = IsVerboseKeyTraceEnabled();
    const LONGLONG keyDownStartCounter = QueryPerfCounterValue();

    *eaten = FALSE;

    const bool ctrlPressed =
        (GetKeyState(VK_CONTROL) & 0x8000) != 0 ||
        (GetKeyState(VK_LCONTROL) & 0x8000) != 0 ||
        (GetKeyState(VK_RCONTROL) & 0x8000) != 0 ||
        (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0 ||
        (GetAsyncKeyState(VK_LCONTROL) & 0x8000) != 0 ||
        (GetAsyncKeyState(VK_RCONTROL) & 0x8000) != 0;
    const bool shiftPressed = (GetKeyState(VK_SHIFT) & 0x8000) != 0;
    const bool altPressed =
        (GetKeyState(VK_MENU) & 0x8000) != 0 ||
        (GetKeyState(VK_LMENU) & 0x8000) != 0 ||
        (GetKeyState(VK_RMENU) & 0x8000) != 0;
    const bool winPressed =
        (GetKeyState(VK_LWIN) & 0x8000) != 0 ||
        (GetKeyState(VK_RWIN) & 0x8000) != 0;

    size_t ctrlSelectionOffset = 0;
    const bool ctrlDigitHotkey = ctrlPressed && TryMapSelectionInputToIndex(key, lParam, ctrlSelectionOffset);
    const bool ctrlShiftModeHotkey = ctrlPressed && shiftPressed && (key == VK_F3 || key == VK_F4);
    const bool ctrlShiftManualHotkey = ctrlPressed && shiftPressed && (key == L'M' || key == L'm');
    const bool ctrlConfigHotkey = ctrlPressed && key == VK_F2;
    const bool genericCtrlEditingShortcut = ctrlPressed &&
        !altPressed &&
        !winPressed &&
        !IsToggleHotkeyPressed(key) &&
        !ctrlDigitHotkey &&
        !(key == VK_DELETE && !compositionCode_.empty()) &&
        !ctrlShiftModeHotkey &&
        !ctrlShiftManualHotkey &&
        !ctrlConfigHotkey;

    if (genericCtrlEditingShortcut) {
        *eaten = FALSE;
        return S_OK;
    }

    if (altPressed || winPressed) {
        *eaten = FALSE;
        return S_OK;
    }

    const bool mayNeedRuntimeInit = chineseMode_ &&
        (IsAlphaKey(key) ||
         !compositionCode_.empty() ||
         key == VK_BACK ||
         key == VK_DELETE ||
         key == VK_SPACE ||
         key == VK_RETURN ||
         key == VK_UP ||
         key == VK_DOWN ||
         key == VK_PRIOR ||
         key == VK_NEXT ||
         key == VK_OEM_MINUS ||
         key == VK_SUBTRACT ||
         key == VK_OEM_PLUS ||
         key == VK_ADD ||
         key == VK_OEM_COMMA ||
         key == VK_OEM_PERIOD);
    if (mayNeedRuntimeInit) {
        EnsureRuntimeReady();
    }

    if (verboseKeyTrace &&
        (IsAlphaKey(key) ||
         (key >= L'0' && key <= L'9') ||
         (key >= VK_NUMPAD0 && key <= VK_NUMPAD9) ||
         key == VK_SPACE ||
         key == VK_RETURN ||
         key == VK_F9 ||
         wParam == VK_PROCESSKEY)) {
        wchar_t keyLog[200] = {};
        swprintf_s(keyLog, L"OnKeyDown raw=0x%04X key=0x%04X chinese=%d codeLen=%u", static_cast<unsigned int>(wParam), static_cast<unsigned int>(key), chineseMode_ ? 1 : 0, static_cast<unsigned int>(compositionCode_.size()));
        Trace(keyLog);
    }

    if (leftShiftTogglePending_ && !(key == VK_SHIFT || key == VK_LSHIFT || key == VK_RSHIFT)) {
        leftShiftTogglePending_ = false;
        leftShiftToggleDownTick_ = 0;
    }

    if (key == VK_SHIFT || key == VK_LSHIFT || key == VK_RSHIFT) {
        const UINT scanCode = (static_cast<UINT>(lParam) >> 16) & 0xFF;
        const bool leftDown =
            (GetKeyState(VK_LSHIFT) & 0x8000) != 0 ||
            (GetAsyncKeyState(VK_LSHIFT) & 0x8000) != 0;
        const bool rightDown =
            (GetKeyState(VK_RSHIFT) & 0x8000) != 0 ||
            (GetAsyncKeyState(VK_RSHIFT) & 0x8000) != 0;
        if (verboseKeyTrace) {
            wchar_t shiftLog[220] = {};
            swprintf_s(
                shiftLog,
                L"shift key event raw=0x%04X key=0x%04X scan=0x%02X left=%d right=%d ctrl=%d alt=%d win=%d",
                static_cast<unsigned int>(wParam),
                static_cast<unsigned int>(key),
                static_cast<unsigned int>(scanCode),
                leftDown ? 1 : 0,
                rightDown ? 1 : 0,
                ctrlPressed ? 1 : 0,
                altPressed ? 1 : 0,
                winPressed ? 1 : 0);
            Trace(shiftLog);
        }
    }

    // Defer Left Shift toggle to key-up. KeyDown stays uneaten so Shift combos keep their native keyboard state.
    if (IsLeftShiftToggleKey(key, lParam) && !ctrlPressed && !altPressed && !winPressed) {
        const bool wasDown = ((static_cast<UINT>(lParam) >> 30) & 0x1U) != 0;
        if (!wasDown) {
            leftShiftTogglePending_ = true;
            leftShiftToggleDownTick_ = GetTickCount64();
        }
        *eaten = FALSE;
        return S_OK;
    }

    if (IsToggleHotkeyPressed(key)) {
        chineseMode_ = !chineseMode_;
        if (!chineseMode_) {
            ClearComposition();
            candidateWindow_.Hide();
        } else {
            UpdateCandidateWindow();
        }
        *eaten = TRUE;
        Trace(std::wstring(L"mode=") + (chineseMode_ ? L"ZH" : L"EN"));
        return S_OK;
    }

    if (key == VK_F10) {
        fullShapeMode_ = !fullShapeMode_;
        if (!compositionCode_.empty()) {
            UpdateCandidateWindow();
        }
        *eaten = TRUE;
        Trace(std::wstring(L"shape=") + (fullShapeMode_ ? L"FULL" : L"HALF"));
        return S_OK;
    }

    if (key == VK_F2 && ctrlPressed) {
        const bool launched = LaunchConfigExecutable();
        *eaten = TRUE;
        Trace(std::wstring(L"hotkey ctrl+f2 open config=") + (launched ? L"1" : L"0"));
        return S_OK;
    }

    if (ctrlPressed && shiftPressed && (key == VK_F3 || key == VK_F4)) {
        DictionaryProfile targetProfile = DictionaryProfile::ZhengmaLarge;
        if (key == VK_F3) {
            targetProfile = DictionaryProfile::ZhengmaLargePinyin;
        }

        SwitchDictionaryProfile(targetProfile);
        *eaten = TRUE;
        return S_OK;
    }

    if (ctrlPressed && shiftPressed && (key == L'M' || key == L'm')) {
        if (PromoteSelectedCandidateToManualEntry()) {
            MessageBeep(MB_OK);
        } else {
            MessageBeep(MB_ICONWARNING);
        }
        *eaten = TRUE;
        return S_OK;
    }

    if (key == VK_F2 || key == VK_APPS) {
        wchar_t menuLog[160] = {};
        swprintf_s(menuLog, L"hotkey menu key=0x%04X ctrl=%d", static_cast<unsigned int>(key), ctrlPressed ? 1 : 0);
        Trace(menuLog);
        const bool shown = ShowStatusMenu();
        if (key == VK_F2 && !ctrlPressed && !shown) {
            const bool launched = LaunchConfigExecutable();
            Trace(std::wstring(L"hotkey f2 fallback open config=") + (launched ? L"1" : L"0"));
        }
        *eaten = TRUE;
        return S_OK;
    }

    if (!chineseMode_) {
        if (!fullShapeMode_) {
            *eaten = FALSE;
            return S_OK;
        }

        if (key == VK_BACK || key == VK_DELETE) {
            *eaten = FALSE;
            return S_OK;
        }

        if (CommitAsciiKey(context, wParam, lParam)) {
            *eaten = TRUE;
        }
        return S_OK;
    }

    auto commitRawCompositionThenKey = [&](WPARAM trailingKey, LPARAM trailingLParam) {
        const std::wstring rawCode = compositionCode_;
        if (!CommitText(context, rawCode)) {
            if (verboseKeyTrace) {
                Trace(L"commit(raw fallback) failed text=" + rawCode);
            }
            return false;
        }

        ClearComposition();
        candidateWindow_.Hide();
        bool committedTrailing = false;
        if ((trailingKey >= L'0' && trailingKey <= L'9') || (trailingKey >= VK_NUMPAD0 && trailingKey <= VK_NUMPAD9)) {
            committedTrailing = CommitAsciiKey(context, trailingKey, trailingLParam);
        }
        if (verboseKeyTrace) {
            Trace(L"commit(raw fallback)=" + rawCode);
        }
        return committedTrailing || true;
    };

    size_t selectionOffset = 0;
    if (ctrlPressed && TryMapSelectionInputToIndex(key, lParam, selectionOffset) && !compositionCode_.empty()) {
        const size_t pageStart = pageIndex_ * pageSize_;
        PinCandidateByGlobalIndex(pageStart + selectionOffset);
        *eaten = TRUE;
        return S_OK;
    }

    if (ctrlPressed && key == VK_DELETE && !compositionCode_.empty()) {
        const size_t pageCandidateCount = GetCurrentPageCandidateCount();
        if (pageCandidateCount > 0) {
            const size_t safeIndexInPage = std::min(selectedIndexInPage_, pageCandidateCount - 1);
            BlockCandidateByGlobalIndex(pageIndex_ * pageSize_ + safeIndexInPage);
        }
        *eaten = TRUE;
        return S_OK;
    }

    const bool minusPageUpKey = IsMinusPageUpKey(key, wParam, lParam);
    const bool plusPageDownKey = IsPlusPageDownKey(key, wParam, lParam);
    if (!minusPageUpKey && !plusPageDownKey) {
        pageBoundaryDirection_ = 0;
        pageBoundaryHitCount_ = 0;
    }

    std::wstring punctuationText;
    if (!compositionCode_.empty() && (minusPageUpKey || plusPageDownKey)) {
        if (plusPageDownKey && !candidatesFullyExpanded_) {
            const size_t requestedPage = pageIndex_ + 1;
            RefreshCandidates(true);
            const size_t expandedTotalPages = GetTotalPages();
            if (expandedTotalPages > 0 && requestedPage < expandedTotalPages) {
                pageIndex_ = requestedPage;
                selectedIndexInPage_ = FindPreferredSelectionIndexForPage(pageIndex_);
                UpdateCandidateWindow();
                *eaten = TRUE;
                return S_OK;
            }
        }

        const size_t totalPages = GetTotalPages();
        const bool canPageUp = pageIndex_ > 0;
        const bool canPageDown = totalPages > 0 && pageIndex_ + 1 < totalPages;
        const bool canMovePage = minusPageUpKey ? canPageUp : canPageDown;

        if (canMovePage) {
            if (minusPageUpKey) {
                pageIndex_ -= 1;
            } else {
                pageIndex_ += 1;
            }
            selectedIndexInPage_ = FindPreferredSelectionIndexForPage(pageIndex_);
            UpdateCandidateWindow();
            pageBoundaryDirection_ = 0;
            pageBoundaryHitCount_ = 0;
            *eaten = TRUE;
            TraceLatencySample(L"keydown-page code=" + compositionCode_, keyDownStartCounter, QueryPerfCounterValue());
            return S_OK;
        }

        const int boundaryDirection = minusPageUpKey ? -1 : 1;
        if (pageBoundaryDirection_ != boundaryDirection) {
            pageBoundaryDirection_ = boundaryDirection;
            pageBoundaryHitCount_ = 1;
            *eaten = TRUE;
            return S_OK;
        }

        pageBoundaryHitCount_ += 1;
        if (pageBoundaryHitCount_ < 2) {
            *eaten = TRUE;
            return S_OK;
        }

        pageBoundaryDirection_ = 0;
        pageBoundaryHitCount_ = 0;

        std::wstring boundaryPunctuation;
        const bool hasBoundaryPunctuation = TryBuildPunctuationCommitText(
            wParam,
            lParam,
            chinesePunctuation_,
            smartSymbolPairs_,
            nextSingleQuoteOpen_,
            nextDoubleQuoteOpen_,
            boundaryPunctuation);

        bool committed = false;
        const size_t pageCandidateCount = GetCurrentPageCandidateCount();
        if (pageCandidateCount > 0) {
            const size_t safeIndexInPage = std::min(selectedIndexInPage_, pageCandidateCount - 1);
            const size_t globalIndex = pageIndex_ * pageSize_ + safeIndexInPage;
            committed = CommitCandidateByGlobalIndex(context, globalIndex, 1);
        } else {
            const std::wstring rawCode = compositionCode_;
            if (CommitText(context, rawCode)) {
                ClearComposition();
                candidateWindow_.Hide();
                committed = true;
            }
        }

        if (committed && hasBoundaryPunctuation) {
            CommitText(context, boundaryPunctuation);
        }

        *eaten = TRUE;
        TraceLatencySample(L"keydown-page-boundary code=" + compositionCode_, keyDownStartCounter, QueryPerfCounterValue());
        return S_OK;
    }

    const bool digitProducesSymbol =
        (key >= L'0' && key <= L'9') &&
        [&]() {
            wchar_t input = 0;
            if (!TryGetTypedChar(wParam, lParam, input)) {
                return false;
            }
            return std::iswalpha(static_cast<wint_t>(input)) == 0 &&
                   std::iswdigit(static_cast<wint_t>(input)) == 0 &&
                   std::iswspace(static_cast<wint_t>(input)) == 0;
        }();

    if (IsPunctuationVirtualKey(key) || digitProducesSymbol) {
        const bool hasPunctuation = TryBuildPunctuationCommitText(
            wParam,
            lParam,
            chinesePunctuation_,
            smartSymbolPairs_,
            nextSingleQuoteOpen_,
            nextDoubleQuoteOpen_,
            punctuationText);

        if (!compositionCode_.empty()) {
            bool committed = false;
            if (!allCandidates_.empty()) {
                committed = CommitCandidateByGlobalIndex(context, 0, 1);
            } else {
                const std::wstring rawCode = compositionCode_;
                if (CommitText(context, rawCode)) {
                    ClearComposition();
                    candidateWindow_.Hide();
                    committed = true;
                    if (verboseKeyTrace) {
                        Trace(L"punctuation fallback commit raw=" + rawCode);
                    }
                }
            }

            if (committed && hasPunctuation) {
                CommitText(context, punctuationText);
                if (verboseKeyTrace) {
                    Trace(L"punctuation after composition commit=" + punctuationText);
                }
            }

            *eaten = TRUE;
            return S_OK;
        }

        if (compositionCode_.empty() && hasPunctuation) {
            if (CommitText(context, punctuationText)) {
                *eaten = TRUE;
                if (verboseKeyTrace) {
                    Trace(L"punctuation=" + punctuationText);
                }
            }
            return S_OK;
        }
    }

    if (IsAlphaKey(key)) {
        const bool pinyinQueryMode = dictionaryProfile_ == DictionaryProfile::ZhengmaLargePinyin;
        const auto countExactCandidates = [this]() {
            size_t count = 0;
            for (const CandidateItem& item : allCandidates_) {
                if (item.exactMatch) {
                    ++count;
                }
            }
            return count;
        };

        const bool shouldSuppressAutoCommitForPinyin =
            pinyinQueryMode &&
            compositionCode_.size() >= 4 &&
            countExactCandidates() != 1;

        const bool shouldAutoCommitDuringTyping = !pinyinQueryMode;

        constexpr size_t kAutoCommitCodeLength = 4;
        if (shouldAutoCommitDuringTyping && !shouldSuppressAutoCommitForPinyin && compositionCode_.size() >= kAutoCommitCodeLength) {
            bool autoCommitted = false;
            if (!allCandidates_.empty()) {
                size_t commitIndex = 0;
                bool foundExact = false;
                for (size_t i = 0; i < allCandidates_.size(); ++i) {
                    if (allCandidates_[i].exactMatch) {
                        commitIndex = i;
                        foundExact = true;
                        break;
                    }
                }

                if (foundExact) {
                    autoCommitted = CommitCandidateByGlobalIndex(context, commitIndex, 1);
                }
                if (autoCommitted) {
                    if (verboseKeyTrace) {
                        Trace(L"auto-commit first candidate on 4-code overflow");
                    }
                }
            }

            if (!autoCommitted) {
                const std::wstring rawCode = compositionCode_;
                if (CommitText(context, rawCode)) {
                    ClearComposition();
                    candidateWindow_.Hide();
                    autoCommitted = true;
                    if (verboseKeyTrace) {
                        Trace(L"auto-commit raw code on 4-code overflow=" + rawCode);
                    }
                }
            }
        }

        if (shouldAutoCommitDuringTyping && !shouldSuppressAutoCommitForPinyin && autoCommitUniqueExact_ && !compositionCode_.empty()) {
            size_t uniqueExactIndex = 0;
            if (compositionCode_.size() >= static_cast<size_t>(autoCommitMinCodeLength_) && TryFindUniqueExactCommitCandidateIndex(uniqueExactIndex)) {
                if (CommitCandidateByGlobalIndex(context, uniqueExactIndex, 1)) {
                    if (verboseKeyTrace) {
                        Trace(L"auto-commit exact before continue input");
                    }
                }
            }
        }
        compositionCode_.push_back(ToLowerAlpha(key));
        RefreshCandidates();
        *eaten = TRUE;
        TraceLatencySample(L"keydown-alpha code=" + compositionCode_, keyDownStartCounter, QueryPerfCounterValue());
        if (verboseKeyTrace) {
            Trace(L"code=" + compositionCode_ + L" candidates=" + std::to_wstring(allCandidates_.size()));
        }
        return S_OK;
    }

    if (tabNavigation_ && key == VK_TAB && !compositionCode_.empty()) {
        const size_t pageCandidateCount = GetCurrentPageCandidateCount();
        if (pageCandidateCount > 0) {
            if (shiftPressed) {
                if (selectedIndexInPage_ == 0) {
                    if (pageIndex_ > 0) {
                        pageIndex_ -= 1;
                        const size_t prevPageCandidateCount = GetCurrentPageCandidateCount();
                        selectedIndexInPage_ = prevPageCandidateCount == 0 ? 0 : (prevPageCandidateCount - 1);
                    } else {
                        selectedIndexInPage_ = pageCandidateCount - 1;
                    }
                } else {
                    selectedIndexInPage_ -= 1;
                }
            } else {
                if (selectedIndexInPage_ + 1 < pageCandidateCount) {
                    selectedIndexInPage_ += 1;
                } else {
                    const size_t totalPages = GetTotalPages();
                    if (totalPages > 0 && pageIndex_ + 1 < totalPages) {
                        pageIndex_ += 1;
                        selectedIndexInPage_ = 0;
                    } else {
                        selectedIndexInPage_ = 0;
                    }
                }
            }
            UpdateCandidateWindow();
        }
        *eaten = TRUE;
        return S_OK;
    }

    if ((key == VK_OEM_4 || key == VK_OEM_COMMA || key == VK_PRIOR) && !compositionCode_.empty()) {
        if (pageIndex_ > 0) {
            pageIndex_ -= 1;
            selectedIndexInPage_ = FindPreferredSelectionIndexForPage(pageIndex_);
            UpdateCandidateWindow();
        }
        *eaten = TRUE;
        return S_OK;
    }

    if ((key == VK_OEM_6 || key == VK_OEM_PERIOD || key == VK_NEXT) && !compositionCode_.empty()) {
        if (!candidatesFullyExpanded_) {
            const size_t requestedPage = pageIndex_ + 1;
            RefreshCandidates(true);
            const size_t expandedTotalPages = GetTotalPages();
            if (expandedTotalPages > 0 && requestedPage < expandedTotalPages) {
                pageIndex_ = requestedPage;
                selectedIndexInPage_ = FindPreferredSelectionIndexForPage(pageIndex_);
                UpdateCandidateWindow();
                *eaten = TRUE;
                return S_OK;
            }
        }

        const size_t totalPages = GetTotalPages();
        if (totalPages > 0 && pageIndex_ + 1 < totalPages) {
            pageIndex_ += 1;
            selectedIndexInPage_ = FindPreferredSelectionIndexForPage(pageIndex_);
            UpdateCandidateWindow();
        }
        *eaten = TRUE;
        return S_OK;
    }

    if ((key == VK_UP || key == VK_DOWN) && !compositionCode_.empty()) {
        const size_t pageCandidateCount = GetCurrentPageCandidateCount();
        if (pageCandidateCount > 0) {
            if (key == VK_UP) {
                if (selectedIndexInPage_ == 0) {
                    selectedIndexInPage_ = pageCandidateCount - 1;
                } else {
                    selectedIndexInPage_ -= 1;
                }
            } else {
                selectedIndexInPage_ = (selectedIndexInPage_ + 1) % pageCandidateCount;
            }
            UpdateCandidateWindow();
        }
        *eaten = TRUE;
        return S_OK;
    }

    if (key == VK_BACK && !compositionCode_.empty()) {
        compositionCode_.pop_back();
        if (compositionCode_.empty()) {
            ClearComposition();
            candidateWindow_.Hide();
        } else if (!TryRestoreCachedCandidatesForCode(compositionCode_)) {
            RefreshCandidates();
        }
        *eaten = TRUE;
        TraceLatencySample(L"keydown-backspace code=" + compositionCode_, keyDownStartCounter, QueryPerfCounterValue());
        if (verboseKeyTrace) {
            Trace(L"backspace code=" + compositionCode_);
        }
        return S_OK;
    }

    if (key == VK_ESCAPE && !compositionCode_.empty()) {
        ClearComposition();
        candidateWindow_.Hide();
        *eaten = TRUE;
        if (verboseKeyTrace) {
            Trace(L"composition cleared");
        }
        return S_OK;
    }

    selectionOffset = 0;
    if (TryMapSelectionInputToIndex(key, lParam, selectionOffset) && !compositionCode_.empty()) {
        if (verboseKeyTrace) {
            Trace(L"select: begin");
        }
        const size_t pageCandidateCount = GetCurrentPageCandidateCount();
        if (selectionOffset < pageCandidateCount) {
            const size_t pageStart = pageIndex_ * pageSize_;
            const size_t index = pageStart + selectionOffset;
            CommitCandidateByGlobalIndex(context, index, 1);
        } else {
            commitRawCompositionThenKey(wParam, lParam);
        }
        *eaten = TRUE;
        return S_OK;
    }

    if ((key == VK_SPACE || key == VK_RETURN) && !compositionCode_.empty()) {
        const size_t pageCandidateCount = GetCurrentPageCandidateCount();
        if (pageCandidateCount == 0) {
            const std::wstring rawCode = compositionCode_;
            if (CommitText(context, rawCode)) {
                ClearComposition();
                candidateWindow_.Hide();
                *eaten = TRUE;
                if (verboseKeyTrace) {
                    Trace(L"commit=" + rawCode);
                }
            } else {
                if (verboseKeyTrace) {
                    Trace(L"commit failed text=" + rawCode);
                }
            }
        } else {
            if (key == VK_RETURN) {
                if (enterExactPriority_) {
                    size_t exactIndex = 0;
                    if (TryFindExactCommitCandidateIndex(exactIndex)) {
                        if (CommitCandidateByGlobalIndex(context, exactIndex, 1)) {
                            *eaten = TRUE;
                        }
                    } else {
                        const std::wstring rawCode = compositionCode_;
                        if (CommitText(context, rawCode)) {
                            ClearComposition();
                            candidateWindow_.Hide();
                            *eaten = TRUE;
                            if (verboseKeyTrace) {
                                Trace(L"commit(raw enter)=" + rawCode);
                            }
                        } else {
                            if (verboseKeyTrace) {
                                Trace(L"commit(raw enter) failed text=" + rawCode);
                            }
                        }
                    }
                } else {
                    const size_t safeIndexInPage = std::min(selectedIndexInPage_, pageCandidateCount - 1);
                    const size_t globalIndex = pageIndex_ * pageSize_ + safeIndexInPage;
                    if (CommitCandidateByGlobalIndex(context, globalIndex, 1)) {
                        *eaten = TRUE;
                    }
                }
            } else {
                const size_t safeIndexInPage = std::min(selectedIndexInPage_, pageCandidateCount - 1);
                const size_t globalIndex = pageIndex_ * pageSize_ + safeIndexInPage;
                if (CommitCandidateByGlobalIndex(context, globalIndex, 1)) {
                    *eaten = TRUE;
                }
            }
        }
        return S_OK;
    }

    return S_OK;
}

STDMETHODIMP TextService::OnKeyUp(ITfContext* context, WPARAM wParam, LPARAM lParam, BOOL* eaten) {
    if (eaten == nullptr) {
        return E_INVALIDARG;
    }

    const WPARAM key = NormalizeVirtualKey(wParam, lParam);
    const bool verboseKeyTrace = IsVerboseKeyTraceEnabled();
    const bool ctrlPressed =
        (GetKeyState(VK_CONTROL) & 0x8000) != 0 ||
        (GetKeyState(VK_LCONTROL) & 0x8000) != 0 ||
        (GetKeyState(VK_RCONTROL) & 0x8000) != 0 ||
        (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0 ||
        (GetAsyncKeyState(VK_LCONTROL) & 0x8000) != 0 ||
        (GetAsyncKeyState(VK_RCONTROL) & 0x8000) != 0;
    const bool altPressed =
        (GetKeyState(VK_MENU) & 0x8000) != 0 ||
        (GetKeyState(VK_LMENU) & 0x8000) != 0 ||
        (GetKeyState(VK_RMENU) & 0x8000) != 0;
    const bool winPressed =
        (GetKeyState(VK_LWIN) & 0x8000) != 0 ||
        (GetKeyState(VK_RWIN) & 0x8000) != 0;

    if (IsLeftShiftToggleKey(key, lParam)) {
        constexpr ULONGLONG kLeftShiftToggleTapThresholdMs = 400ULL;
        const ULONGLONG nowTick = GetTickCount64();
        const bool withinTapWindow =
            leftShiftToggleDownTick_ != 0 &&
            nowTick >= leftShiftToggleDownTick_ &&
            (nowTick - leftShiftToggleDownTick_) <= kLeftShiftToggleTapThresholdMs;
        const bool shouldToggle = leftShiftTogglePending_ && withinTapWindow && !ctrlPressed && !altPressed && !winPressed;
        if (verboseKeyTrace) {
            wchar_t shiftUpLog[260] = {};
            swprintf_s(
                shiftUpLog,
                L"OnKeyUp shift raw=0x%04X key=0x%04X pending=%d withinTap=%d shouldToggle=%d chinese=%d fullShape=%d",
                static_cast<unsigned int>(wParam),
                static_cast<unsigned int>(key),
                leftShiftTogglePending_ ? 1 : 0,
                withinTapWindow ? 1 : 0,
                shouldToggle ? 1 : 0,
                chineseMode_ ? 1 : 0,
                fullShapeMode_ ? 1 : 0);
            Trace(shiftUpLog);
        }
        leftShiftTogglePending_ = false;
        leftShiftToggleDownTick_ = 0;
        if (shouldToggle) {
            if (chineseMode_) {
                chineseMode_ = false;
                chinesePunctuation_ = false;
                fullShapeMode_ = false;
                if (!compositionCode_.empty()) {
                    ClearComposition();
                }
                candidateWindow_.Hide();
                Trace(L"left-shift toggle mode=EN punctuation=HALF shape=HALF");
            } else {
                chineseMode_ = true;
                chinesePunctuation_ = true;
                UpdateCandidateWindow();
                Trace(std::wstring(L"left-shift toggle mode=ZH punctuation=") + (chinesePunctuation_ ? L"FULL" : L"HALF") +
                      L" shape=" + (fullShapeMode_ ? L"FULL" : L"HALF"));
            }
            *eaten = TRUE;
            return S_OK;
        }
    }

    *eaten = FALSE;
    return S_OK;
}

STDMETHODIMP TextService::OnPreservedKey(ITfContext*, REFGUID, BOOL* eaten) {
    if (eaten == nullptr) {
        return E_INVALIDARG;
    }

    *eaten = FALSE;
    return S_OK;
}

HRESULT CreateTextServiceClassFactory(REFIID riid, void** ppv) {
    if (ppv == nullptr) {
        return E_INVALIDARG;
    }

    *ppv = nullptr;
    auto* factory = new (std::nothrow) TextServiceClassFactory();
    if (factory == nullptr) {
        return E_OUTOFMEMORY;
    }

    const HRESULT hr = factory->QueryInterface(riid, ppv);
    factory->Release();
    return hr;
}
