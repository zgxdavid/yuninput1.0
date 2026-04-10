#include "CandidateWindow.h"

#include "Globals.h"

#include <algorithm>
#include <cmath>
#include <sstream>

namespace {

constexpr wchar_t kCandidateWindowClass[] = L"yuninput_candidate_window";
constexpr size_t kMaxVisibleRows = 6;
constexpr const wchar_t* kIndexLabels[] = {L"1", L"2", L"3", L"4", L"5", L"6", L"7", L"8", L"9"};
constexpr UINT_PTR kAsyncPollTimerId = 1;
constexpr size_t kMaxCandidateWidthChars = 12;
constexpr int kNormalCharWidthPx = 17;
constexpr int kSelectedCharWidthPx = 19;
constexpr int kCodeColumnWidthPx = 74;
constexpr int kWindowBaseWidthPx = 116;
constexpr int kWindowMinWidthPx = 280;
constexpr int kWindowMaxWidthPx = 430;
constexpr int kPositionSnapThresholdPx = 3;

bool AreDisplayCandidatesEqual(
    const std::vector<CandidateWindow::DisplayCandidate>& left,
    const std::vector<CandidateWindow::DisplayCandidate>& right) {
    if (left.size() != right.size()) {
        return false;
    }

    for (size_t i = 0; i < left.size(); ++i) {
        const auto& l = left[i];
        const auto& r = right[i];
        if (l.text != r.text ||
            l.code != r.code ||
            l.boostedUser != r.boostedUser ||
            l.boostedLearned != r.boostedLearned ||
            l.boostedContext != r.boostedContext ||
            l.fromAutoPhrase != r.fromAutoPhrase ||
            l.fromSessionAutoPhrase != r.fromSessionAutoPhrase ||
            l.consumedLength != r.consumedLength) {
            return false;
        }
    }

    return true;
}

bool GetWorkAreaForPoint(const POINT& point, RECT& outWorkArea) {
    HMONITOR monitor = MonitorFromPoint(point, MONITOR_DEFAULTTONEAREST);
    if (monitor != nullptr) {
        MONITORINFO monitorInfo = {};
        monitorInfo.cbSize = sizeof(monitorInfo);
        if (GetMonitorInfoW(monitor, &monitorInfo)) {
            outWorkArea = monitorInfo.rcWork;
            return true;
        }
    }

    return SystemParametersInfoW(SPI_GETWORKAREA, 0, &outWorkArea, 0) == TRUE;
}

void FillSolidRect(HDC hdc, const RECT& rc, COLORREF color) {
    SetDCBrushColor(hdc, color);
    FillRect(hdc, &rc, static_cast<HBRUSH>(GetStockObject(DC_BRUSH)));
}

void DrawRoundedRect(HDC hdc, const RECT& rc, int ellipseW, int ellipseH, COLORREF fillColor, COLORREF borderColor) {
    HGDIOBJ oldPen = SelectObject(hdc, GetStockObject(DC_PEN));
    HGDIOBJ oldBrush = SelectObject(hdc, GetStockObject(DC_BRUSH));
    SetDCPenColor(hdc, borderColor);
    SetDCBrushColor(hdc, fillColor);
    RoundRect(hdc, rc.left, rc.top, rc.right, rc.bottom, ellipseW, ellipseH);
    SelectObject(hdc, oldBrush);
    SelectObject(hdc, oldPen);
}

void FillVerticalGradient(HDC hdc, const RECT& rc, COLORREF topColor, COLORREF bottomColor) {
    const int height = rc.bottom - rc.top;
    if (height <= 0) {
        return;
    }

    using GradientFillProc = BOOL(WINAPI*)(HDC, PTRIVERTEX, ULONG, PVOID, ULONG, ULONG);
    static GradientFillProc sGradientFill = []() -> GradientFillProc {
        HMODULE mod = LoadLibraryW(L"msimg32.dll");
        if (mod == nullptr) {
            return nullptr;
        }
        return reinterpret_cast<GradientFillProc>(GetProcAddress(mod, "GradientFill"));
    }();

    if (sGradientFill != nullptr) {
        TRIVERTEX vertex[2] = {};
        vertex[0].x = rc.left;
        vertex[0].y = rc.top;
        vertex[0].Red = static_cast<COLOR16>(GetRValue(topColor) << 8);
        vertex[0].Green = static_cast<COLOR16>(GetGValue(topColor) << 8);
        vertex[0].Blue = static_cast<COLOR16>(GetBValue(topColor) << 8);

        vertex[1].x = rc.right;
        vertex[1].y = rc.bottom;
        vertex[1].Red = static_cast<COLOR16>(GetRValue(bottomColor) << 8);
        vertex[1].Green = static_cast<COLOR16>(GetGValue(bottomColor) << 8);
        vertex[1].Blue = static_cast<COLOR16>(GetBValue(bottomColor) << 8);

        GRADIENT_RECT gRect = {0, 1};
        if (sGradientFill(hdc, vertex, 2, &gRect, 1, GRADIENT_FILL_RECT_V)) {
            return;
        }
    }

    const int topR = GetRValue(topColor);
    const int topG = GetGValue(topColor);
    const int topB = GetBValue(topColor);
    const int bottomR = GetRValue(bottomColor);
    const int bottomG = GetGValue(bottomColor);
    const int bottomB = GetBValue(bottomColor);

    HPEN pen = CreatePen(PS_SOLID, 1, topColor);
    HGDIOBJ oldPen = SelectObject(hdc, pen);
    for (int i = 0; i < height; ++i) {
        const int r = topR + ((bottomR - topR) * i) / std::max(1, height - 1);
        const int g = topG + ((bottomG - topG) * i) / std::max(1, height - 1);
        const int b = topB + ((bottomB - topB) * i) / std::max(1, height - 1);
        SetDCPenColor(hdc, RGB(r, g, b));
        MoveToEx(hdc, rc.left, rc.top + i, nullptr);
        LineTo(hdc, rc.right, rc.top + i);
    }
    SelectObject(hdc, oldPen);
    DeleteObject(pen);
}

COLORREF BlendColor(COLORREF left, COLORREF right, double ratio) {
    const double safeRatio = std::max(0.0, std::min(1.0, ratio));
    const int r = static_cast<int>(GetRValue(left) + (GetRValue(right) - GetRValue(left)) * safeRatio);
    const int g = static_cast<int>(GetGValue(left) + (GetGValue(right) - GetGValue(left)) * safeRatio);
    const int b = static_cast<int>(GetBValue(left) + (GetBValue(right) - GetBValue(left)) * safeRatio);
    return RGB(r, g, b);
}

SIZE MeasureText(HDC hdc, HFONT font, const std::wstring& text) {
    SIZE size = {};
    if (text.empty()) {
        return size;
    }

    HGDIOBJ oldFont = SelectObject(hdc, font);
    GetTextExtentPoint32W(hdc, text.c_str(), static_cast<int>(text.size()), &size);
    SelectObject(hdc, oldFont);
    return size;
}

bool IsAutoPhraseCandidate(const CandidateWindow::DisplayCandidate& candidate) {
    return candidate.fromAutoPhrase || candidate.fromSessionAutoPhrase;
}

COLORREF GetCandidateTextColor(const CandidateWindow::DisplayCandidate& candidate, bool selected) {
    const bool learned = candidate.boostedLearned;
    const bool autoPhrase = IsAutoPhraseCandidate(candidate);
    if (learned && autoPhrase) {
        return selected ? RGB(238, 220, 255) : RGB(134, 74, 181);
    }
    if (learned) {
        return selected ? RGB(214, 255, 224) : RGB(48, 130, 74);
    }
    if (autoPhrase) {
        return selected ? RGB(255, 235, 190) : RGB(186, 124, 24);
    }
    return selected ? RGB(255, 252, 245) : RGB(44, 37, 28);
}

}  // namespace

CandidateWindow::CandidateWindow()
    : hwnd_(nullptr),
      pageIndex_(0),
      totalPages_(0),
      totalCandidateCount_(0),
      selectedIndex_(0),
      selectedAbsoluteIndex_(0),
      chineseMode_(true),
      fullShapeMode_(false),
      displayedQualityLevel_(0),
      qualityLevelTick_(0),
      qualityPulseUntilTick_(0),
      hasLastWindowRect_(false),
      lastX_(0),
      lastY_(0),
      lastWidth_(0),
      lastHeight_(0),
      lastMeasuredRowCount_(0),
      lastMeasuredTotalPages_(0),
      lastMeasuredTotalCandidateCount_(0),
      titleFont_(nullptr),
      textFont_(nullptr),
      selectedTextFont_(nullptr),
      smallFont_(nullptr),
      codeFont_(nullptr) {}

CandidateWindow::~CandidateWindow() {
    Destroy();
}

bool CandidateWindow::EnsureCreated() {
    if (hwnd_ != nullptr) {
        return true;
    }

    WNDCLASSW wc = {};
    wc.lpfnWndProc = CandidateWindow::WndProc;
    wc.hInstance = g_moduleHandle;
    wc.lpszClassName = kCandidateWindowClass;
    wc.hCursor = LoadCursorW(nullptr, IDC_ARROW);
    wc.hbrBackground = reinterpret_cast<HBRUSH>(COLOR_WINDOW + 1);

    RegisterClassW(&wc);

    hwnd_ = CreateWindowExW(
        WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE,
        kCandidateWindowClass,
        L"yuninput candidates",
        WS_POPUP,
        120,
        120,
        360,
        220,
        nullptr,
        nullptr,
        g_moduleHandle,
        this);

    EnsureFonts();
    return hwnd_ != nullptr;
}

bool CandidateWindow::EnsureFonts() {
    if (titleFont_ != nullptr &&
        textFont_ != nullptr &&
        selectedTextFont_ != nullptr &&
        smallFont_ != nullptr &&
        codeFont_ != nullptr) {
        return true;
    }

    if (titleFont_ == nullptr) {
        titleFont_ = CreateFontW(-17, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
            OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Microsoft YaHei UI");
    }
    if (textFont_ == nullptr) {
        textFont_ = CreateFontW(-18, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
            OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Microsoft YaHei UI");
    }
    if (selectedTextFont_ == nullptr) {
        selectedTextFont_ = CreateFontW(-20, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
            OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Microsoft YaHei UI");
    }
    if (smallFont_ == nullptr) {
        smallFont_ = CreateFontW(-14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
            OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Microsoft YaHei UI");
    }
    if (codeFont_ == nullptr) {
        codeFont_ = CreateFontW(-15, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
            OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, FIXED_PITCH | FF_MODERN, L"Consolas");
    }

    return titleFont_ != nullptr &&
           textFont_ != nullptr &&
           selectedTextFont_ != nullptr &&
           smallFont_ != nullptr &&
           codeFont_ != nullptr;
}

void CandidateWindow::ReleaseFonts() {
    if (titleFont_ != nullptr) {
        DeleteObject(titleFont_);
        titleFont_ = nullptr;
    }
    if (textFont_ != nullptr) {
        DeleteObject(textFont_);
        textFont_ = nullptr;
    }
    if (selectedTextFont_ != nullptr) {
        DeleteObject(selectedTextFont_);
        selectedTextFont_ = nullptr;
    }
    if (smallFont_ != nullptr) {
        DeleteObject(smallFont_);
        smallFont_ = nullptr;
    }
    if (codeFont_ != nullptr) {
        DeleteObject(codeFont_);
        codeFont_ = nullptr;
    }
}

void CandidateWindow::Destroy() {
    if (hwnd_ != nullptr) {
        DestroyWindow(hwnd_);
        hwnd_ = nullptr;
    }
    ReleaseFonts();
    hasLastWindowRect_ = false;
    lastMeasuredRowCount_ = 0;
    lastMeasuredTotalPages_ = 0;
    lastMeasuredTotalCandidateCount_ = 0;
}

void CandidateWindow::Update(
    const std::wstring& code,
    const std::vector<DisplayCandidate>& candidates,
    size_t pageIndex,
    size_t totalPages,
    size_t totalCandidateCount,
    size_t selectedIndex,
    size_t selectedAbsoluteIndex,
    bool chineseMode,
    bool fullShapeMode,
    const POINT* anchorScreenPos) {
    if (!EnsureCreated()) {
        return;
    }

    const bool cheapStateChanged =
        code_ != code ||
        pageIndex_ != pageIndex ||
        totalPages_ != totalPages ||
        totalCandidateCount_ != totalCandidateCount ||
        selectedIndex_ != selectedIndex ||
        selectedAbsoluteIndex_ != selectedAbsoluteIndex ||
        chineseMode_ != chineseMode ||
        fullShapeMode_ != fullShapeMode;

    const bool candidatesChanged = cheapStateChanged ? true : !AreDisplayCandidatesEqual(candidates_, candidates);
    const bool contentChanged = cheapStateChanged || candidatesChanged;

    if (contentChanged) {
        code_ = code;
        candidates_ = candidates;
        pageIndex_ = pageIndex;
        totalPages_ = totalPages;
        totalCandidateCount_ = totalCandidateCount;
        selectedIndex_ = selectedIndex;
        selectedAbsoluteIndex_ = selectedAbsoluteIndex;
        chineseMode_ = chineseMode;
        fullShapeMode_ = fullShapeMode;
    }

    const size_t visibleRowCount = std::max<size_t>(1, std::min(candidates_.size(), kMaxVisibleRows));
    const int rowHeight = 42;
    int width = kWindowMinWidthPx;
    const int height = 94 + static_cast<int>(visibleRowCount) * rowHeight + 42;
    int measuredWidth = 312;
    for (size_t i = 0; i < visibleRowCount; ++i) {
        const auto& candidate = candidates_[i];
        const size_t visibleChars = std::min(candidate.text.size(), kMaxCandidateWidthChars);
        const int charWidth = std::max(kNormalCharWidthPx, kSelectedCharWidthPx);
        int mainWidth = static_cast<int>(visibleChars) * charWidth;
        if (candidate.text.size() > kMaxCandidateWidthChars) {
            mainWidth += 10;
        }

        const int rowWidth = kWindowBaseWidthPx + mainWidth + kCodeColumnWidthPx;
        measuredWidth = std::max(measuredWidth, rowWidth);
    }

    width = std::max(kWindowMinWidthPx, std::min(kWindowMaxWidthPx, measuredWidth));

    int x = 120;
    int y = 120;
    POINT anchorPoint = {120, 120};
    bool hasAnchor = false;
    if (anchorScreenPos != nullptr) {
        anchorPoint = *anchorScreenPos;
        hasAnchor = true;
    } else {
        hasAnchor = GetCursorPos(&anchorPoint) == TRUE;
    }

    if (hasAnchor) {
        x = anchorPoint.x + 10;
        y = anchorPoint.y + 18;
    }

    RECT workArea = {};
    if (GetWorkAreaForPoint(anchorPoint, workArea)) {
        if (hasAnchor) {
            const int belowY = anchorPoint.y + 18;
            const int aboveY = anchorPoint.y - height - 10;
            if (belowY + height <= workArea.bottom || aboveY < workArea.top) {
                y = belowY;
            } else {
                y = aboveY;
            }
        }

        const int maxX = std::max(static_cast<int>(workArea.left), static_cast<int>(workArea.right) - width);
        const int maxY = std::max(static_cast<int>(workArea.top), static_cast<int>(workArea.bottom) - height);
        x = std::max(static_cast<int>(workArea.left), std::min(x, maxX));
        y = std::max(static_cast<int>(workArea.top), std::min(y, maxY));
    }

    if (hasLastWindowRect_ &&
        std::abs(x - lastX_) <= kPositionSnapThresholdPx &&
        std::abs(y - lastY_) <= kPositionSnapThresholdPx) {
        x = lastX_;
        y = lastY_;
    }

    const bool sizeChanged =
        !hasLastWindowRect_ ||
        width != lastWidth_ ||
        height != lastHeight_;
    const bool moved = !hasLastWindowRect_ || x != lastX_ || y != lastY_;

    lastMeasuredRowCount_ = visibleRowCount;
    lastMeasuredTotalPages_ = totalPages;
    lastMeasuredTotalCandidateCount_ = totalCandidateCount;
    lastX_ = x;
    lastY_ = y;
    lastWidth_ = width;
    lastHeight_ = height;
    hasLastWindowRect_ = true;

    if (moved || sizeChanged || !IsVisible()) {
        SetWindowPos(hwnd_, HWND_TOP, x, y, width, height, SWP_SHOWWINDOW | SWP_NOACTIVATE);
    }

    if (contentChanged || sizeChanged || !IsVisible()) {
        InvalidateRect(hwnd_, nullptr, FALSE);
    }
}

bool CandidateWindow::IsVisible() const {
    return hwnd_ != nullptr && IsWindowVisible(hwnd_) == TRUE;
}

void CandidateWindow::SetAsyncPollCallback(std::function<void()> callback) {
    asyncPollCallback_ = std::move(callback);
}

void CandidateWindow::ScheduleAsyncPoll(UINT delayMs) {
    if (hwnd_ == nullptr) {
        return;
    }

    SetTimer(hwnd_, kAsyncPollTimerId, std::max<UINT>(1, delayMs), nullptr);
}

void CandidateWindow::CancelAsyncPoll() {
    if (hwnd_ == nullptr) {
        return;
    }

    KillTimer(hwnd_, kAsyncPollTimerId);
}

bool CandidateWindow::Hide() {
    if (!IsVisible()) {
        return false;
    }

    CancelAsyncPoll();
    ShowWindow(hwnd_, SW_HIDE);
    return true;
}

void CandidateWindow::OnPaint() {
    EnsureFonts();
    PAINTSTRUCT ps = {};
    HDC hdc = BeginPaint(hwnd_, &ps);

    RECT rc = {};
    GetClientRect(hwnd_, &rc);

    const int width = rc.right - rc.left;
    const int height = rc.bottom - rc.top;
    if (width <= 0 || height <= 0) {
        EndPaint(hwnd_, &ps);
        return;
    }

    HDC memoryDc = CreateCompatibleDC(hdc);
    HBITMAP bitmap = nullptr;
    HGDIOBJ oldBitmap = nullptr;
    if (memoryDc != nullptr) {
        bitmap = CreateCompatibleBitmap(hdc, width, height);
        if (bitmap != nullptr) {
            oldBitmap = SelectObject(memoryDc, bitmap);
            PaintContent(memoryDc, rc);
            BitBlt(hdc, 0, 0, width, height, memoryDc, 0, 0, SRCCOPY);
            SelectObject(memoryDc, oldBitmap);
        } else {
            DeleteDC(memoryDc);
            memoryDc = nullptr;
        }
    }

    if (memoryDc == nullptr) {
        PaintContent(hdc, rc);
    }

    if (bitmap != nullptr) {
        DeleteObject(bitmap);
    }
    if (memoryDc != nullptr) {
        DeleteDC(memoryDc);
    }

    EndPaint(hwnd_, &ps);
}

void CandidateWindow::PaintContent(HDC hdc, const RECT& rc) const {

    const COLORREF bgTop = RGB(246, 239, 226);
    const COLORREF bgBottom = RGB(236, 244, 248);
    const COLORREF panelBg = RGB(255, 252, 246);
    const COLORREF border = RGB(153, 130, 101);
    const COLORREF softBorder = RGB(210, 191, 167);
    const COLORREF textMain = RGB(44, 37, 28);
    const COLORREF textMeta = RGB(102, 91, 79);
    const COLORREF titleAccent = RGB(78, 117, 165);
    const COLORREF selectedBgTop = RGB(66, 118, 186);
    const COLORREF selectedBgBottom = RGB(43, 82, 136);
    const COLORREF selectedBorder = RGB(27, 58, 102);
    const int codeColumnWidth = 74;

    FillVerticalGradient(hdc, rc, bgTop, bgBottom);

    RECT panelRc = rc;
    panelRc.left += 4;
    panelRc.top += 4;
    panelRc.right -= 4;
    panelRc.bottom -= 4;
    DrawRoundedRect(hdc, panelRc, 18, 18, panelBg, border);

    RECT titleStrip = panelRc;
    titleStrip.left += 1;
    titleStrip.top += 1;
    titleStrip.right -= 1;
    titleStrip.bottom = titleStrip.top + 34;
    FillVerticalGradient(hdc, titleStrip, RGB(255, 249, 238), RGB(241, 231, 214));

    SetBkMode(hdc, TRANSPARENT);
    HGDIOBJ oldFont = SelectObject(hdc, titleFont_ != nullptr ? titleFont_ : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT)));

    RECT headerRc = panelRc;
    headerRc.left += 14;
    headerRc.right -= 14;
    headerRc.top += 7;
    headerRc.bottom = headerRc.top + 22;
    SetTextColor(hdc, titleAccent);
    const std::wstring title = L"Yuninput \u5019\u9009";
    DrawTextW(hdc, title.c_str(), static_cast<int>(title.size()), &headerRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

    wchar_t pageText[96] = {};
    swprintf_s(
        pageText,
        L"\u7b2c %zu/%zu \u9875  \u5171 %zu \u9879",
        totalPages_ == 0 ? static_cast<size_t>(0) : (pageIndex_ + 1),
        totalPages_,
        totalCandidateCount_);
    SetTextColor(hdc, textMeta);
    DrawTextW(hdc, pageText, -1, &headerRc, DT_RIGHT | DT_VCENTER | DT_SINGLELINE);

    SelectObject(hdc, codeFont_ != nullptr ? codeFont_ : titleFont_);
    RECT codeRc = panelRc;
    codeRc.left += 14;
    codeRc.right -= 14;
    codeRc.top = headerRc.bottom + 8;
    codeRc.bottom = codeRc.top + 28;
    DrawRoundedRect(hdc, codeRc, 14, 14, RGB(247, 242, 232), softBorder);
    const std::wstring codeLine = std::wstring(L"Code  ") + code_;
    RECT codeTextRc = codeRc;
    codeTextRc.left += 12;
    SetTextColor(hdc, textMain);
    DrawTextW(hdc, codeLine.c_str(), static_cast<int>(codeLine.size()), &codeTextRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

    int y = codeRc.bottom + 10;
    const int rowHeight = 42;
    const int rowLeft = panelRc.left + 14;
    const int rowRight = panelRc.right - 14;
    const size_t rowCount = std::min(candidates_.size(), kMaxVisibleRows);

    if (rowCount == 0) {
        RECT emptyRc = panelRc;
        emptyRc.left += 14;
        emptyRc.right -= 14;
        emptyRc.top = y;
        emptyRc.bottom = emptyRc.top + 24;
        SetTextColor(hdc, textMeta);
        const std::wstring empty = L"\u6ca1\u6709\u53ef\u7528\u5019\u9009";
        DrawTextW(hdc, empty.c_str(), static_cast<int>(empty.size()), &emptyRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
    } else {
        for (size_t i = 0; i < rowCount; ++i) {
            const DisplayCandidate& candidate = candidates_[i];
            const bool isSelected = i == selectedIndex_;
            RECT rowRc = {rowLeft, y, rowRight, y + rowHeight};
            if (isSelected) {
                DrawRoundedRect(hdc, rowRc, 14, 14, selectedBgBottom, selectedBorder);
                RECT selectedHighlight = rowRc;
                selectedHighlight.left += 1;
                selectedHighlight.top += 1;
                selectedHighlight.right -= 1;
                selectedHighlight.bottom = selectedHighlight.top + (rowHeight / 2);
                FillVerticalGradient(hdc, selectedHighlight, selectedBgTop, selectedBgBottom);
            } else {
                const COLORREF rowFill = BlendColor(panelBg, RGB(240, 234, 224), (i % 2 == 0) ? 0.14 : 0.28);
                DrawRoundedRect(hdc, rowRc, 14, 14, rowFill, softBorder);
            }

            RECT indexRc = rowRc;
            indexRc.left += 8;
            indexRc.top += 8;
            indexRc.right = indexRc.left + 28;
            indexRc.bottom = indexRc.top + 24;
            DrawRoundedRect(
                hdc,
                indexRc,
                12,
                12,
                isSelected ? RGB(255, 248, 232) : RGB(233, 223, 208),
                isSelected ? RGB(255, 214, 140) : RGB(191, 176, 153));
            SetTextColor(hdc, RGB(83, 62, 35));
            SelectObject(hdc, titleFont_ != nullptr ? titleFont_ : oldFont);
            const wchar_t* indexText = (i < _countof(kIndexLabels)) ? kIndexLabels[i] : L"-";
            DrawTextW(hdc, indexText, -1, &indexRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

            RECT textRc = rowRc;
            textRc.left = indexRc.right + 10;
            textRc.top += 9;
            textRc.right = rowRc.right - codeColumnWidth - 16;
            textRc.bottom = textRc.top + 22;
            SetTextColor(hdc, GetCandidateTextColor(candidate, isSelected));
            SelectObject(hdc, isSelected ? selectedTextFont_ : textFont_);
            DrawTextW(hdc, candidate.text.c_str(), static_cast<int>(candidate.text.size()), &textRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

            RECT codeHintRc = rowRc;
            codeHintRc.left = rowRc.right - codeColumnWidth;
            codeHintRc.right -= 10;
            codeHintRc.top += 10;
            codeHintRc.bottom = codeHintRc.top + 18;
            SetTextColor(hdc, isSelected ? RGB(216, 230, 255) : titleAccent);
            SelectObject(hdc, codeFont_ != nullptr ? codeFont_ : smallFont_);
            DrawTextW(hdc, candidate.code.c_str(), static_cast<int>(candidate.code.size()), &codeHintRc, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

            RECT metaRc = rowRc;
            metaRc.left = rowRc.right - codeColumnWidth;
            metaRc.right -= 10;
            metaRc.top = rowRc.top + 25;
            metaRc.bottom = metaRc.top + 14;
            SetTextColor(hdc, isSelected ? RGB(238, 243, 249) : textMeta);
            SelectObject(hdc, smallFont_ != nullptr ? smallFont_ : textFont_);
            if (candidate.consumedLength > 0 && candidate.consumedLength < code_.size()) {
                std::wstringstream meta;
                meta << L"\u4f59\u7801 " << (code_.size() - candidate.consumedLength);
                const std::wstring metaText = meta.str();
                DrawTextW(hdc, metaText.c_str(), static_cast<int>(metaText.size()), &metaRc, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
            }

            y += rowHeight + 6;
        }
    }

    SelectObject(hdc, smallFont_ != nullptr ? smallFont_ : textFont_);
    RECT footerRc = panelRc;
    footerRc.left += 14;
    footerRc.right -= 14;
    footerRc.bottom -= 10;
    footerRc.top = footerRc.bottom - 20;
    SetTextColor(hdc, textMeta);
    const std::wstring footer = L"Space/Enter \u4e0a\u5c4f  \u6570\u5b57\u76f4\u9009  \u4e0a\u4e0b\u79fb\u52a8  PgUp/PgDn \u7ffb\u9875";
    DrawTextW(hdc, footer.c_str(), static_cast<int>(footer.size()), &footerRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

    SelectObject(hdc, oldFont);
}

LRESULT CALLBACK CandidateWindow::WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    CandidateWindow* self = reinterpret_cast<CandidateWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    if (msg == WM_NCCREATE) {
        CREATESTRUCTW* createStruct = reinterpret_cast<CREATESTRUCTW*>(lParam);
        self = static_cast<CandidateWindow*>(createStruct->lpCreateParams);
        if (self != nullptr) {
            self->hwnd_ = hwnd;
            SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        }
    }

    if (self == nullptr) {
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }

    switch (msg) {
    case WM_PAINT:
        self->OnPaint();
        return 0;
    case WM_ERASEBKGND:
        return 1;
    case WM_TIMER:
        if (wParam == kAsyncPollTimerId) {
            KillTimer(hwnd, kAsyncPollTimerId);
            if (self->asyncPollCallback_) {
                self->asyncPollCallback_();
            }
            return 0;
        }
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    case WM_NCDESTROY:
        KillTimer(hwnd, kAsyncPollTimerId);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, 0);
        self->hwnd_ = nullptr;
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    default:
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }
}
