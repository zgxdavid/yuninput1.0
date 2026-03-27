#include "TextService.h"

#include "Globals.h"

#include <Windows.h>
#include <shellapi.h>

#include <algorithm>
#include <cwctype>
#include <filesystem>
#include <fstream>
#include <new>
#include <regex>
#include <sstream>

namespace {

constexpr const wchar_t* kBuildMarker = L"cw-r3-20260327-keytrace-v1";
constexpr ULONGLONG kAnchorReuseWindowMs = 2500ULL;

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

    const int required = MultiByteToWideChar(CP_UTF8, 0, input.c_str(), static_cast<int>(input.size()), nullptr, 0);
    if (required <= 0) {
        return L"";
    }

    std::wstring output(static_cast<size_t>(required), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, input.c_str(), static_cast<int>(input.size()), output.data(), required);
    return output;
}

std::wstring GetRuntimeLogPath() {
    wchar_t localAppData[MAX_PATH] = {};
    const DWORD len = GetEnvironmentVariableW(L"LOCALAPPDATA", localAppData, MAX_PATH);
    if (len == 0 || len >= MAX_PATH) {
        return L"C:\\Windows\\Temp\\yuninput_runtime.log";
    }

    return std::wstring(localAppData, len) + L"\\yuninput\\runtime.log";
}

void AppendRuntimeLog(const std::wstring& message) {
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
    GetLocalTime(&now);
    wchar_t prefix[64] = {};
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
           key == VK_OEM_PERIOD ||
           key == VK_OEM_PLUS;
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
    std::wifstream input(path);
    if (!input.is_open()) {
        return L"";
    }

    input.imbue(std::locale(".UTF-8"));
    std::wstring content;
    std::wstring line;
    while (std::getline(input, line)) {
        content.append(line);
        content.push_back(L'\n');
    }
    return content;
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
      pageSize_(kDefaultPageSize),
      toggleHotkey_(ToggleHotkey::F9),
      pageIndex_(0),
    selectedIndexInPage_(0),
        emptyCandidateAlerted_(false),
        hasRecentAnchor_(false),
        lastAnchor_{0, 0},
                lastAnchorTick_(0),
                autoPhraseSelectedStreak_(0),
                autoPhraseSelectedTick_(0) {
    InterlockedIncrement(&g_objectCount);
}

TextService::~TextService() {
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

    LoadConfiguredDictionaries();

    if (EnsureUserDataDirectory(userDataDir_)) {
        userDictPath_ = userDataDir_ + L"\\user_dict.txt";
        autoPhraseDictPath_ = userDataDir_ + L"\\auto_phrase_dict.txt";
        userFreqPath_ = userDataDir_ + L"\\user_freq.txt";
        blockedEntriesPath_ = userDataDir_ + L"\\blocked_entries.txt";
        contextAssocPath_ = userDataDir_ + L"\\context_assoc.txt";
        contextAssocBlacklistPath_ = userDataDir_ + L"\\context_assoc_blacklist.txt";
        manualPhraseReviewPath_ = userDataDir_ + L"\\manual_phrase_review.txt";
        engine_.LoadUserDictionaryFromFile(userDictPath_);
        engine_.LoadAutoPhraseDictionaryFromFile(autoPhraseDictPath_);
        engine_.LoadFrequencyFromFile(userFreqPath_);
        engine_.LoadBlockedEntriesFromFile(blockedEntriesPath_);
        LoadContextAssociationFromFile(contextAssocPath_);
        LoadContextAssociationBlacklistFromFile(contextAssocBlacklistPath_);
    }

    candidateWindow_.EnsureCreated();

    Trace(std::wstring(L"TextService activated marker=") + kBuildMarker);

    return S_OK;
}

STDMETHODIMP TextService::Deactivate() {
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
    candidateWindow_.Destroy();
    Trace(L"TextService deactivated");
    return S_OK;
}

void TextService::RefreshCandidates() {
    allCandidates_.clear();
    if (compositionCode_.empty()) {
        pageIndex_ = 0;
        selectedIndexInPage_ = 0;
        emptyCandidateAlerted_ = false;
        UpdateCandidateWindow();
        return;
    }

    const bool pinyinFallbackMode = dictionaryProfile_ == DictionaryProfile::ZhengmaLargePinyin;

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

    const std::vector<CompositionEngine::Entry> queried = engine_.QueryCandidateEntries(compositionCode_, 200);
    allCandidates_.reserve(queried.size() + 32);

    auto mergeCandidate = [this](
                              const std::wstring& text,
                              const std::wstring& displayCode,
                              const std::wstring& commitCode,
                              std::uint64_t contextScore,
                              bool exactMatch,
                              bool boostedUser,
                              bool boostedLearned,
                              bool boostedContext,
                              bool fromAutoPhrase,
                              bool fromSystemDict,
                              size_t consumedLength) {
        if (text.empty()) {
            return;
        }

        for (CandidateItem& existing : allCandidates_) {
            if (existing.text != text) {
                continue;
            }

            existing.boostedUser = existing.boostedUser || boostedUser;
            existing.boostedLearned = existing.boostedLearned || boostedLearned;
            existing.boostedContext = existing.boostedContext || boostedContext;
            existing.exactMatch = existing.exactMatch || exactMatch;
            existing.fromAutoPhrase = existing.fromAutoPhrase || fromAutoPhrase;
            existing.fromSystemDict = existing.fromSystemDict || fromSystemDict;
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

            if ((boostedContext && !existing.boostedContext) ||
                consumedLength > existing.consumedLength ||
                existing.commitCode.empty()) {
                existing.commitCode = commitCode;
                existing.consumedLength = consumedLength;
            }
            return;
        }

        CandidateItem candidate;
        candidate.text = text;
        candidate.code = displayCode;
        candidate.commitCode = commitCode;
        candidate.contextScore = contextScore;
        candidate.exactMatch = exactMatch;
        candidate.boostedUser = boostedUser;
        candidate.boostedLearned = boostedLearned;
        candidate.boostedContext = boostedContext;
        candidate.fromAutoPhrase = fromAutoPhrase;
        candidate.fromSystemDict = fromSystemDict;
        candidate.consumedLength = consumedLength;
        allCandidates_.push_back(std::move(candidate));
    };

    bool hasExactCurrentCode = false;
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
            item.code == compositionCode_,
            item.isUser && !item.isAutoPhrase,
            item.isLearned,
            false,
            item.isAutoPhrase,
            !item.isUser,
            compositionCode_.size());
    }

    if (pinyinFallbackMode && allCandidates_.empty() && compositionCode_.size() > 2) {
        for (size_t prefixLength = compositionCode_.size() - 1; prefixLength >= 2; --prefixLength) {
            const std::wstring prefixCode = compositionCode_.substr(0, prefixLength);
            const std::vector<CompositionEngine::Entry> prefixMatches = engine_.QueryCandidateEntries(prefixCode, 80);
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
                    false,
                    item.isUser && !item.isAutoPhrase,
                    item.isLearned,
                    false,
                    item.isAutoPhrase,
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

    if (!hasExactCurrentCode && compositionCode_.size() > 1 && !pinyinFallbackMode) {
        for (size_t prefixLength = compositionCode_.size() - 1; prefixLength >= 1; --prefixLength) {
            const std::wstring prefixCode = compositionCode_.substr(0, prefixLength);
            const std::vector<CompositionEngine::Entry> prefixMatches = engine_.QueryCandidateEntries(prefixCode, 24);
            for (const auto& item : prefixMatches) {
                if (item.code != prefixCode) {
                    continue;
                }

                mergeCandidate(
                    item.text,
                    item.code,
                    prefixCode,
                    0,
                    false,
                    item.isUser && !item.isAutoPhrase,
                    item.isLearned,
                    false,
                    item.isAutoPhrase,
                    !item.isUser,
                    prefixLength);
            }

            if (prefixLength == 1) {
                break;
            }
        }
    }

    const ULONGLONG now = GetTickCount64();
    if (!recentCommits_.empty() && !pinyinFallbackMode) {
        const CommitHistoryItem& previous = recentCommits_.back();
        if ((now - previous.tick) <= 8000ULL && !previous.code.empty() && !previous.text.empty()) {
            const std::wstring phraseCode = previous.code + compositionCode_;
            const std::vector<CompositionEngine::Entry> phraseMatches = engine_.QueryCandidateEntries(phraseCode, 64);
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
                    false,
                    item.isUser && !item.isAutoPhrase,
                    true,
                    true,
                    item.isAutoPhrase,
                    !item.isUser,
                    compositionCode_.size());
            }
        }
    }

    if (contextAssociationEnabled_ && !recentCommits_.empty()) {
        const CommitHistoryItem& previous = recentCommits_.back();
        if ((now - previous.tick) <= 15000ULL && !previous.text.empty()) {
            for (CandidateItem& candidate : allCandidates_) {
                const std::uint64_t contextScore = QueryContextAssociationScore(previous.text, candidate.text);
                if (contextScore == 0) {
                    continue;
                }

                candidate.contextScore = std::max(candidate.contextScore, contextScore);
                candidate.boostedContext = true;
            }
        }
    }

    const bool repeatBoostActive =
        autoPhraseSelectedStreak_ > 3 &&
        !lastAutoPhraseSelectedKey_.empty() &&
        now >= autoPhraseSelectedTick_ &&
        (now - autoPhraseSelectedTick_) <= 60000ULL;
    for (CandidateItem& candidate : allCandidates_) {
        candidate.boostedAutoRepeat = false;
        if (!repeatBoostActive || !candidate.fromAutoPhrase) {
            continue;
        }

        const std::wstring candidateKey =
            (candidate.commitCode.empty() ? compositionCode_ : candidate.commitCode) + L"\t" + candidate.text;
        candidate.boostedAutoRepeat = candidateKey == lastAutoPhraseSelectedKey_;
    }

    const size_t queryCodeLength = compositionCode_.size();
    std::stable_sort(
        allCandidates_.begin(),
        allCandidates_.end(),
        [pinyinFallbackMode, queryCodeLength](const CandidateItem& left, const CandidateItem& right) {
            if (pinyinFallbackMode) {
                const bool leftSingleChar = left.text.size() == 1;
                const bool rightSingleChar = right.text.size() == 1;
                if (leftSingleChar != rightSingleChar) {
                    return leftSingleChar > rightSingleChar;
                }
            }
            if (queryCodeLength <= 2) {
                const bool leftSingleChar = left.text.size() == 1;
                const bool rightSingleChar = right.text.size() == 1;
                const bool leftFirstLevelSingleChar = leftSingleChar && left.code.size() <= 1 && left.commitCode.size() <= 1;
                const bool rightFirstLevelSingleChar = rightSingleChar && right.code.size() <= 1 && right.commitCode.size() <= 1;
                if (leftFirstLevelSingleChar != rightFirstLevelSingleChar) {
                    return leftFirstLevelSingleChar > rightFirstLevelSingleChar;
                }

                const bool leftExactTwoCodePhrase =
                    left.text.size() == 2 &&
                    left.code.size() == 2 &&
                    left.commitCode.size() == 2 &&
                    left.fromSystemDict &&
                    !left.fromAutoPhrase;
                const bool rightExactTwoCodePhrase =
                    right.text.size() == 2 &&
                    right.code.size() == 2 &&
                    right.commitCode.size() == 2 &&
                    right.fromSystemDict &&
                    !right.fromAutoPhrase;
                if (leftExactTwoCodePhrase != rightExactTwoCodePhrase) {
                    return leftExactTwoCodePhrase > rightExactTwoCodePhrase;
                }

                const bool leftSecondLevelSingleChar = leftSingleChar && left.code.size() == 2 && left.commitCode.size() == 2;
                const bool rightSecondLevelSingleChar = rightSingleChar && right.code.size() == 2 && right.commitCode.size() == 2;
                if (leftSecondLevelSingleChar != rightSecondLevelSingleChar) {
                    return leftSecondLevelSingleChar > rightSecondLevelSingleChar;
                }
            }
            if (left.exactMatch != right.exactMatch) {
                return left.exactMatch > right.exactMatch;
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
            const bool leftAutoOnly = left.fromAutoPhrase && !left.fromSystemDict && !left.boostedUser;
            const bool rightAutoOnly = right.fromAutoPhrase && !right.fromSystemDict && !right.boostedUser;
            if (leftAutoOnly != rightAutoOnly) {
                return leftAutoOnly < rightAutoOnly;
            }
            if (left.consumedLength != right.consumedLength) {
                return left.consumedLength > right.consumedLength;
            }
            if (left.boostedUser != right.boostedUser) {
                return left.boostedUser > right.boostedUser;
            }
            if (left.boostedLearned != right.boostedLearned) {
                return left.boostedLearned > right.boostedLearned;
            }
            return false;
        });

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

    UpdateCandidateWindow();
}

const wchar_t* TextService::GetDictionaryProfileName(DictionaryProfile profile) {
    switch (profile) {
    case DictionaryProfile::ZhengmaLargePinyin:
        return L"zhengma-large-pinyin";
    case DictionaryProfile::ZhengmaLarge:
    default:
        return L"zhengma-large";
    }
}

bool TextService::ReloadActiveDictionaries() {
    const bool loaded = LoadConfiguredDictionaries();

    if (!userDictPath_.empty()) {
        engine_.LoadUserDictionaryFromFile(userDictPath_);
    }
    if (!autoPhraseDictPath_.empty()) {
        engine_.LoadAutoPhraseDictionaryFromFile(autoPhraseDictPath_);
    }
    if (!userFreqPath_.empty()) {
        engine_.LoadFrequencyFromFile(userFreqPath_);
    }
    if (!blockedEntriesPath_.empty()) {
        engine_.LoadBlockedEntriesFromFile(blockedEntriesPath_);
    }

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

    EnsureSingleCharZhengmaCodeHintsLoaded(dataDir);

    std::filesystem::path profileDictPath;
    std::wstring profileName = GetDictionaryProfileName(dictionaryProfile_);
    switch (dictionaryProfile_) {
    case DictionaryProfile::ZhengmaLargePinyin:
        profileDictPath = dataDir / "zhengma-pinyin.dict";
        break;
    case DictionaryProfile::ZhengmaLarge:
    default:
        profileDictPath = dataDir / "zhengma-large.dict";
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

    const bool loadedPackagedUser = engine_.LoadUserDictionaryFromFile(packagedUserDictPath.wstring());
    Trace(
        L"Dictionary profile=" + profileName +
        L" profile_loaded=" + (loadedProfileDict ? L"1" : L"0") +
        L" fallback_loaded=" + (loadedFallback ? L"1" : L"0") +
        L" packaged_user=" + (loadedPackagedUser ? L"1" : L"0") +
        (dictionaryProfile_ == DictionaryProfile::ZhengmaLargePinyin ? L" mode=pinyin-single-char-with-zhengma-hint" : L""));

    return loadedProfileDict || loadedFallback;
}

void TextService::EnsureSingleCharZhengmaCodeHintsLoaded(const std::filesystem::path& dataDir) {
    singleCharZhengmaCodeHints_.clear();

    const std::filesystem::path zhengmaPath = dataDir / "zhengma-large.dict";
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

size_t TextService::FindPreferredSelectionIndexForPage(size_t pageIndex) const {
    const std::vector<CandidateWindow::DisplayCandidate> pageCandidates = GetCurrentPageCandidates();
    if (pageCandidates.empty()) {
        return 0;
    }

    auto scoreCandidate = [this](const CandidateWindow::DisplayCandidate& candidate) {
        int score = 0;
        const bool exactFull = !compositionCode_.empty() && candidate.consumedLength >= compositionCode_.size();
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
        score += static_cast<int>(candidate.consumedLength);
        return score;
    };

    size_t bestIndex = 0;
    int bestScore = scoreCandidate(pageCandidates[0]);
    for (size_t i = 1; i < pageCandidates.size(); ++i) {
        const int score = scoreCandidate(pageCandidates[i]);
        if (score > bestScore) {
            bestScore = score;
            bestIndex = i;
        }
    }

    (void)pageIndex;
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
    std::ifstream input(filePath, std::ios::in | std::ios::binary);
    if (!input) {
        return false;
    }

    contextAssociationScores_.clear();

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
    std::ifstream input(filePath, std::ios::in | std::ios::binary);
    if (!input) {
        return false;
    }

    contextAssociationBlacklist_.clear();

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

    if (contextAssociationBlacklist_.find(MakeContextAssociationKey(prevText, nextText)) != contextAssociationBlacklist_.end()) {
        return 0;
    }

    const auto it = contextAssociationScores_.find(MakeContextAssociationKey(prevText, nextText));
    if (it == contextAssociationScores_.end()) {
        return 0;
    }

    return it->second;
}

void TextService::UpdateCandidateWindow() {
    if (compositionCode_.empty()) {
        candidateWindow_.Hide();
        return;
    }

    POINT anchor = {};
    const POINT* anchorPtr = nullptr;
    bool hasFocusedContext = false;

    ITfContext* focusedContext = nullptr;
    if (TryGetFocusedContext(threadMgr_, &focusedContext)) {
        hasFocusedContext = true;
        if (TryGetCaretScreenPointFromContext(focusedContext, clientId_, anchor)) {
            anchorPtr = &anchor;
            hasRecentAnchor_ = true;
            lastAnchor_ = anchor;
            lastAnchorTick_ = GetTickCount64();
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
        lastAnchorTick_ = GetTickCount64();
    }

    candidateWindow_.Update(
        compositionCode_,
        GetCurrentPageCandidates(),
        pageIndex_,
        GetTotalPages(),
        allCandidates_.size(),
        selectedIndexInPage_,
        pageIndex_ * pageSize_ + selectedIndexInPage_,
        chineseMode_,
        fullShapeMode_,
        anchorPtr);
}

std::vector<CandidateWindow::DisplayCandidate> TextService::GetCurrentPageCandidates() const {
    std::vector<CandidateWindow::DisplayCandidate> page;
    if (allCandidates_.empty()) {
        return page;
    }

    const size_t start = pageIndex_ * pageSize_;
    if (start >= allCandidates_.size()) {
        return page;
    }

    const size_t end = std::min(start + pageSize_, allCandidates_.size());
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

        const size_t consumedLength = std::min(allCandidates_[i].consumedLength, compositionCode_.size());
        if (consumedLength < compositionCode_.size()) {
            const std::wstring remaining = compositionCode_.substr(consumedLength);
            if (!remaining.empty()) {
                if (!candidate.code.empty()) {
                    candidate.code += L"  ";
                }
                candidate.code += L"未完:" + remaining;
            }
        }
        candidate.boostedUser = allCandidates_[i].boostedUser;
        candidate.boostedLearned = allCandidates_[i].boostedLearned;
        candidate.boostedContext = allCandidates_[i].boostedContext;
        candidate.consumedLength = allCandidates_[i].consumedLength;
        page.push_back(std::move(candidate));
    }
    return page;
}

size_t TextService::GetTotalPages() const {
    if (allCandidates_.empty()) {
        return 0;
    }

    return (allCandidates_.size() + pageSize_ - 1) / pageSize_;
}

void TextService::ClearComposition() {
    compositionCode_.clear();
    allCandidates_.clear();
    pageIndex_ = 0;
    selectedIndexInPage_ = 0;
    emptyCandidateAlerted_ = false;
}

void TextService::LearnPhraseFromRecentCommits(const std::wstring& committedCode, const std::wstring& committedText) {
    if (committedCode.empty() || committedText.empty()) {
        return;
    }

    const ULONGLONG now = GetTickCount64();

    for (auto it = pendingPhraseStats_.begin(); it != pendingPhraseStats_.end();) {
        if ((now - it->second.lastTick) > 60000ULL) {
            it = pendingPhraseStats_.erase(it);
        } else {
            ++it;
        }
    }

    while (pendingPhraseStats_.size() > 256) {
        auto oldest = pendingPhraseStats_.begin();
        for (auto it = pendingPhraseStats_.begin(); it != pendingPhraseStats_.end(); ++it) {
            if (it->second.lastTick < oldest->second.lastTick) {
                oldest = it;
            }
        }
        pendingPhraseStats_.erase(oldest);
    }

    while (!recentCommits_.empty() && (now - recentCommits_.front().tick) > 60000ULL) {
        recentCommits_.pop_front();
    }

    const auto isLikelyPinyinSingleChar = [this](const std::wstring& code, const std::wstring& text) {
        return dictionaryProfile_ == DictionaryProfile::ZhengmaLargePinyin &&
               text.size() == 1 &&
               code.size() >= 2;
    };

    const bool currentIsPinyinSingleChar = isLikelyPinyinSingleChar(committedCode, committedText);

    if (!recentCommits_.empty()) {
        const CommitHistoryItem& prev = recentCommits_.back();
        const bool closeEnough = (now - prev.tick) <= 8000ULL;
        if (closeEnough) {
            RecordContextAssociation(prev.text, committedText, 1);
            if (!contextAssocPath_.empty()) {
                SaveContextAssociationToFile(contextAssocPath_);
            }
        }
        constexpr size_t kMaxLearnSpan = 6;
        std::vector<CommitHistoryItem> window;
        window.reserve(recentCommits_.size() + 1);
        for (const auto& item : recentCommits_) {
            if (item.text.empty() || item.code.empty()) {
                continue;
            }
            if ((now - item.tick) > 60000ULL) {
                continue;
            }
            if (isLikelyPinyinSingleChar(item.code, item.text)) {
                continue;
            }
            window.push_back(item);
        }

        CommitHistoryItem current;
        current.code = committedCode;
        current.text = committedText;
        current.tick = now;
        if (!current.text.empty() && !current.code.empty() && !currentIsPinyinSingleChar) {
            window.push_back(std::move(current));
        }

        bool learnedAnyPhrase = false;
        if (window.size() >= 2) {
            const size_t end = window.size() - 1;
            const size_t minStart = end >= (kMaxLearnSpan - 1) ? (end - (kMaxLearnSpan - 1)) : 0;
            for (size_t start = end; start > minStart; --start) {
                const size_t phraseStart = start - 1;

                std::wstring phraseText;
                std::wstring fallbackCode;
                phraseText.reserve(16);
                fallbackCode.reserve(24);

                for (size_t idx = phraseStart; idx <= end; ++idx) {
                    phraseText += window[idx].text;
                    fallbackCode += window[idx].code;
                }

                if (phraseText.size() < 2 || phraseText.size() > 8) {
                    continue;
                }

                std::wstring phraseCode;
                if (!engine_.TryBuildPhraseCode(phraseText, phraseCode)) {
                    phraseCode = fallbackCode;
                }
                if (phraseCode.empty()) {
                    continue;
                }
                if (phraseCode.size() > 20) {
                    phraseCode.resize(20);
                }

                const std::wstring phraseKey = phraseCode + L"\t" + phraseText;
                PendingPhraseStat& stat = pendingPhraseStats_[phraseKey];
                stat.hitCount += 1;
                stat.lastTick = now;

                const std::uint32_t requiredHits = phraseText.size() <= 4 ? 1U : 2U;

                if (stat.hitCount >= requiredHits) {
                    const bool added = engine_.AddAutoPhraseEntry(phraseCode, phraseText);
                    if (added) {
                        engine_.RecordCommit(phraseCode, phraseText, 2);
                        AppendPhraseReviewEntry(phraseCode, phraseText, L"auto");
                        learnedAnyPhrase = true;
                        Trace(
                            L"learn phrase(auto table contiguous)=" + phraseText +
                            L" code=" + phraseCode +
                            L" hits=" + std::to_wstring(stat.hitCount) +
                            L" threshold=" + std::to_wstring(requiredHits));
                    }
                    pendingPhraseStats_.erase(phraseKey);
                }
            }
        }

        if (learnedAnyPhrase && !autoPhraseDictPath_.empty()) {
            engine_.SaveAutoPhraseDictionaryToFile(autoPhraseDictPath_);
        }
    }

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
    if (candidate.fromAutoPhrase) {
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
    if (!userFreqPath_.empty()) {
        engine_.SaveFrequencyToFile(userFreqPath_);
    }

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
    if (!userDictPath_.empty()) {
        engine_.SaveUserDictionaryToFile(userDictPath_);
    }

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
        engine_.SaveBlockedEntriesToFile(blockedEntriesPath_);
    }
    if (!userDictPath_.empty()) {
        engine_.SaveUserDictionaryToFile(userDictPath_);
    }
    if (!userFreqPath_.empty()) {
        engine_.SaveFrequencyToFile(userFreqPath_);
    }

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

    return CommitText(context, std::wstring(1, c));
}

bool TextService::PromoteSelectedCandidateToManualEntry() {
    if (compositionCode_.empty()) {
        return false;
    }

    const std::vector<CandidateWindow::DisplayCandidate> pageCandidates = GetCurrentPageCandidates();
    if (pageCandidates.empty()) {
        return false;
    }

    const size_t safeIndexInPage = std::min(selectedIndexInPage_, pageCandidates.size() - 1);
    const size_t globalIndex = pageIndex_ * pageSize_ + safeIndexInPage;
    if (globalIndex >= allCandidates_.size()) {
        return false;
    }

    const CandidateItem& candidate = allCandidates_[globalIndex];
    const std::wstring manualCode = candidate.commitCode.empty() ? compositionCode_ : candidate.commitCode;
    if (manualCode.empty()) {
        return false;
    }

    engine_.PinEntry(manualCode, candidate.text);
    if (!userDictPath_.empty()) {
        engine_.SaveUserDictionaryToFile(userDictPath_);
    }
    AppendPhraseReviewEntry(manualCode, candidate.text, L"manual");
    RefreshCandidates();
    Trace(L"manual phrase code=" + manualCode + L" text=" + candidate.text);
    return true;
}

void TextService::AppendPhraseReviewEntry(const std::wstring& code, const std::wstring& text, const wchar_t* sourceTag) const {
    if (manualPhraseReviewPath_.empty() || code.empty() || text.empty() || sourceTag == nullptr) {
        return;
    }

    std::ofstream output(manualPhraseReviewPath_, std::ios::out | std::ios::binary | std::ios::app);
    if (!output) {
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
    output.write(line.data(), static_cast<std::streamsize>(line.size()));
}

bool TextService::CommitText(ITfContext* context, const std::wstring& text) {
    (void)context;
    if (text.empty()) {
        return false;
    }

    const bool sent = SendUnicodeTextWithInput(text);
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
        if (!compositionCode_.empty()) {
            ClearComposition();
        }
        candidateWindow_.Hide();
        Trace(L"OnSetFocus(BOOL): hide candidate window");
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
        if (!compositionCode_.empty()) {
            ClearComposition();
        }
        candidateWindow_.Hide();
        Trace(L"OnUninitDocumentMgr: hide candidate window");
    }
    return S_OK;
}

STDMETHODIMP TextService::OnSetFocus(ITfDocumentMgr* documentMgrFocus, ITfDocumentMgr* documentMgrPrevFocus) {
    if (documentMgrFocus != documentMgrPrevFocus) {
        if (!compositionCode_.empty()) {
            ClearComposition();
        }
        candidateWindow_.Hide();
        Trace(L"OnSetFocus(DocumentMgr): hide candidate window");
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
            if (key == VK_BACK || key == VK_DELETE || key == VK_OEM_MINUS || key == VK_OEM_PLUS || IsLeftShiftToggleKey(key, lParam)) {
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

    *eaten = (IsAlphaKey(key) ||
              (key >= L'0' && key <= L'9') ||
              (key >= L'1' && key <= L'9') ||
              (key >= VK_NUMPAD0 && key <= VK_NUMPAD9) ||
              (tabNavigation_ && key == VK_TAB) ||
              (key == VK_BACK && hasComposition) ||
              (key == VK_DELETE && hasComposition) ||
              key == VK_SPACE ||
              key == VK_RETURN ||
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
              key == VK_OEM_MINUS ||
              key == VK_OEM_PLUS ||
              key == VK_OEM_COMMA ||
              key == VK_OEM_PERIOD ||
              key == VK_OEM_4 ||
                  key == VK_OEM_5 ||
                  key == VK_OEM_7 ||
              key == VK_OEM_6)
                 ? TRUE
                 : FALSE;
    if (IsAlphaKey(key) ||
        (key >= L'0' && key <= L'9') ||
        (key >= VK_NUMPAD0 && key <= VK_NUMPAD9) ||
        key == VK_SPACE ||
        key == VK_RETURN ||
        key == VK_F9 ||
        wParam == VK_PROCESSKEY) {
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

    *eaten = FALSE;

    if (IsAlphaKey(key) ||
        (key >= L'0' && key <= L'9') ||
        (key >= VK_NUMPAD0 && key <= VK_NUMPAD9) ||
        key == VK_SPACE ||
        key == VK_RETURN ||
        key == VK_F9 ||
        wParam == VK_PROCESSKEY) {
        wchar_t keyLog[200] = {};
        swprintf_s(keyLog, L"OnKeyDown raw=0x%04X key=0x%04X chinese=%d codeLen=%u", static_cast<unsigned int>(wParam), static_cast<unsigned int>(key), chineseMode_ ? 1 : 0, static_cast<unsigned int>(compositionCode_.size()));
        Trace(keyLog);
    }

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
            Trace(L"commit(raw fallback) failed text=" + rawCode);
            return false;
        }

        ClearComposition();
        candidateWindow_.Hide();
        bool committedTrailing = false;
        if ((trailingKey >= L'0' && trailingKey <= L'9') || (trailingKey >= VK_NUMPAD0 && trailingKey <= VK_NUMPAD9)) {
            committedTrailing = CommitAsciiKey(context, trailingKey, trailingLParam);
        }
        Trace(L"commit(raw fallback)=" + rawCode);
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
        const std::vector<CandidateWindow::DisplayCandidate> pageCandidates = GetCurrentPageCandidates();
        if (!pageCandidates.empty()) {
            const size_t safeIndexInPage = std::min(selectedIndexInPage_, pageCandidates.size() - 1);
            BlockCandidateByGlobalIndex(pageIndex_ * pageSize_ + safeIndexInPage);
        }
        *eaten = TRUE;
        return S_OK;
    }

    std::wstring punctuationText;
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
                    Trace(L"punctuation fallback commit raw=" + rawCode);
                }
            }

            if (committed && hasPunctuation) {
                CommitText(context, punctuationText);
                Trace(L"punctuation after composition commit=" + punctuationText);
            }

            *eaten = TRUE;
            return S_OK;
        }

        if (compositionCode_.empty() && hasPunctuation) {
            if (CommitText(context, punctuationText)) {
                *eaten = TRUE;
                Trace(L"punctuation=" + punctuationText);
            }
            return S_OK;
        }
    }

    if (IsAlphaKey(key)) {
        const bool pinyinQueryMode = dictionaryProfile_ == DictionaryProfile::ZhengmaLargePinyin;
        const auto countExactCandidates = [this]() {
            size_t count = 0;
            const size_t currentCodeLength = compositionCode_.size();
            for (const CandidateItem& item : allCandidates_) {
                if (item.consumedLength >= currentCodeLength) {
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
                const size_t currentCodeLength = compositionCode_.size();
                size_t commitIndex = 0;
                for (size_t i = 0; i < allCandidates_.size(); ++i) {
                    if (allCandidates_[i].consumedLength >= currentCodeLength) {
                        commitIndex = i;
                        break;
                    }
                }

                autoCommitted = CommitCandidateByGlobalIndex(context, commitIndex, 1);
                if (autoCommitted) {
                    Trace(L"auto-commit first candidate on 4-code overflow");
                }
            }

            if (!autoCommitted) {
                const std::wstring rawCode = compositionCode_;
                if (CommitText(context, rawCode)) {
                    ClearComposition();
                    candidateWindow_.Hide();
                    autoCommitted = true;
                    Trace(L"auto-commit raw code on 4-code overflow=" + rawCode);
                }
            }
        }

        if (shouldAutoCommitDuringTyping && !shouldSuppressAutoCommitForPinyin && autoCommitUniqueExact_ && !compositionCode_.empty()) {
            size_t uniqueExactIndex = 0;
            if (compositionCode_.size() >= static_cast<size_t>(autoCommitMinCodeLength_) && TryFindUniqueExactCommitCandidateIndex(uniqueExactIndex)) {
                if (CommitCandidateByGlobalIndex(context, uniqueExactIndex, 2)) {
                    Trace(L"auto-commit exact before continue input");
                }
            }
        }
        compositionCode_.push_back(ToLowerAlpha(key));
        RefreshCandidates();
        *eaten = TRUE;
        Trace(L"code=" + compositionCode_ + L" candidates=" + std::to_wstring(allCandidates_.size()));
        return S_OK;
    }

    if (tabNavigation_ && key == VK_TAB && !compositionCode_.empty()) {
        const std::vector<CandidateWindow::DisplayCandidate> pageCandidates = GetCurrentPageCandidates();
        if (!pageCandidates.empty()) {
            if (shiftPressed) {
                if (selectedIndexInPage_ == 0) {
                    if (pageIndex_ > 0) {
                        pageIndex_ -= 1;
                        const std::vector<CandidateWindow::DisplayCandidate> prevPageCandidates = GetCurrentPageCandidates();
                        selectedIndexInPage_ = prevPageCandidates.empty() ? 0 : (prevPageCandidates.size() - 1);
                    } else {
                        selectedIndexInPage_ = pageCandidates.size() - 1;
                    }
                } else {
                    selectedIndexInPage_ -= 1;
                }
            } else {
                if (selectedIndexInPage_ + 1 < pageCandidates.size()) {
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

    if ((key == VK_OEM_MINUS || key == VK_OEM_4 || key == VK_OEM_COMMA || key == VK_PRIOR) && !compositionCode_.empty()) {
        if (pageIndex_ > 0) {
            pageIndex_ -= 1;
            selectedIndexInPage_ = FindPreferredSelectionIndexForPage(pageIndex_);
            UpdateCandidateWindow();
        }
        *eaten = TRUE;
        return S_OK;
    }

    if ((key == VK_OEM_PLUS || key == VK_OEM_6 || key == VK_OEM_PERIOD || key == VK_NEXT) && !compositionCode_.empty()) {
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
        const std::vector<CandidateWindow::DisplayCandidate> pageCandidates = GetCurrentPageCandidates();
        if (!pageCandidates.empty()) {
            if (key == VK_UP) {
                if (selectedIndexInPage_ == 0) {
                    selectedIndexInPage_ = pageCandidates.size() - 1;
                } else {
                    selectedIndexInPage_ -= 1;
                }
            } else {
                selectedIndexInPage_ = (selectedIndexInPage_ + 1) % pageCandidates.size();
            }
            UpdateCandidateWindow();
        }
        *eaten = TRUE;
        return S_OK;
    }

    if (key == VK_BACK && !compositionCode_.empty()) {
        compositionCode_.pop_back();
        RefreshCandidates();
        *eaten = TRUE;
        Trace(L"backspace code=" + compositionCode_);
        return S_OK;
    }

    if (key == VK_ESCAPE && !compositionCode_.empty()) {
        ClearComposition();
        candidateWindow_.Hide();
        *eaten = TRUE;
        Trace(L"composition cleared");
        return S_OK;
    }

    selectionOffset = 0;
    if (TryMapSelectionInputToIndex(key, lParam, selectionOffset) && !compositionCode_.empty()) {
        Trace(L"select: begin");
        const std::vector<CandidateWindow::DisplayCandidate> pageCandidates = GetCurrentPageCandidates();
        if (selectionOffset < pageCandidates.size()) {
            const size_t pageStart = pageIndex_ * pageSize_;
            const size_t index = pageStart + selectionOffset;
            CommitCandidateByGlobalIndex(context, index, 3);
        } else {
            commitRawCompositionThenKey(wParam, lParam);
        }
        *eaten = TRUE;
        return S_OK;
    }

    if ((key == VK_SPACE || key == VK_RETURN) && !compositionCode_.empty()) {
        const std::vector<CandidateWindow::DisplayCandidate> pageCandidates = GetCurrentPageCandidates();
        if (pageCandidates.empty()) {
            const std::wstring rawCode = compositionCode_;
            if (CommitText(context, rawCode)) {
                ClearComposition();
                candidateWindow_.Hide();
                *eaten = TRUE;
                Trace(L"commit=" + rawCode);
            } else {
                Trace(L"commit failed text=" + rawCode);
            }
        } else {
            if (key == VK_RETURN) {
                if (enterExactPriority_) {
                    size_t exactIndex = 0;
                    if (TryFindExactCommitCandidateIndex(exactIndex)) {
                        if (CommitCandidateByGlobalIndex(context, exactIndex, 2)) {
                            *eaten = TRUE;
                        }
                    } else {
                        const std::wstring rawCode = compositionCode_;
                        if (CommitText(context, rawCode)) {
                            ClearComposition();
                            candidateWindow_.Hide();
                            *eaten = TRUE;
                            Trace(L"commit(raw enter)=" + rawCode);
                        } else {
                            Trace(L"commit(raw enter) failed text=" + rawCode);
                        }
                    }
                } else {
                    const size_t safeIndexInPage = std::min(selectedIndexInPage_, pageCandidates.size() - 1);
                    const size_t globalIndex = pageIndex_ * pageSize_ + safeIndexInPage;
                    if (CommitCandidateByGlobalIndex(context, globalIndex, 2)) {
                        *eaten = TRUE;
                    }
                }
            } else {
                const size_t safeIndexInPage = std::min(selectedIndexInPage_, pageCandidates.size() - 1);
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
