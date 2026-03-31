#include "CandidateWindow.h"

#include "Globals.h"

#include <algorithm>
#include <cmath>
#include <sstream>

namespace {

constexpr wchar_t kCandidateWindowClass[] = L"yuninput_candidate_window";
constexpr size_t kMaxVisibleRows = 6;
constexpr const wchar_t* kIndexLabels[] = {L"1", L"2", L"3", L"4", L"5", L"6", L"7", L"8", L"9"};

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
            l.consumedLength != r.consumedLength) {
            return false;
        }
    }

    return true;
}

int ApproachInt(int current, int target, int minStep, double ratio) {
    const int delta = target - current;
    if (delta == 0) {
        return current;
    }

    int step = static_cast<int>(std::abs(static_cast<double>(delta)) * ratio);
    if (step < minStep) {
        step = minStep;
    }
    if (step > std::abs(delta)) {
        step = std::abs(delta);
    }

    return current + ((delta > 0) ? step : -step);
}

int ApproachIntElastic(
    int current,
    int target,
    int minStep,
    double growRatio,
    double shrinkRatio,
    int snapDistance) {
    const int delta = target - current;
    const int absDelta = std::abs(delta);
    if (absDelta <= snapDistance) {
        return target;
    }

    const double ratio = (delta > 0) ? growRatio : shrinkRatio;
    return ApproachInt(current, target, minStep, ratio);
}

int ApproachIntSpring(
    int current,
    int target,
    int minStep,
    double ratio,
    int snapDistance,
    int overshootThreshold,
    int maxOvershootPx) {
    const int delta = target - current;
    const int absDelta = std::abs(delta);
    if (absDelta <= snapDistance) {
        return target;
    }

    int step = static_cast<int>(static_cast<double>(absDelta) * ratio);
    if (step < minStep) {
        step = minStep;
    }

    if (absDelta <= overshootThreshold) {
        const int overshoot = std::min(maxOvershootPx, std::max(2, absDelta / 3));
        step = absDelta + overshoot;
    }
    else if (step > absDelta) {
        step = absDelta;
    }

    return current + ((delta > 0) ? step : -step);
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

void DrawRectBorder(HDC hdc, const RECT& rc, COLORREF borderColor) {
    HGDIOBJ oldPen = SelectObject(hdc, GetStockObject(DC_PEN));
    HGDIOBJ oldBrush = SelectObject(hdc, GetStockObject(HOLLOW_BRUSH));
    SetDCPenColor(hdc, borderColor);
    Rectangle(hdc, rc.left, rc.top, rc.right, rc.bottom);
    SelectObject(hdc, oldBrush);
    SelectObject(hdc, oldPen);
}

void FillVerticalGradient(HDC hdc, const RECT& rc, COLORREF topColor, COLORREF bottomColor) {
    const int height = rc.bottom - rc.top;
    if (height <= 0) {
        return;
    }

    using GradientFillProc = BOOL (WINAPI*)(
        HDC,
        PTRIVERTEX,
        ULONG,
        PVOID,
        ULONG,
        ULONG);
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
        vertex[0].Alpha = 0;

        vertex[1].x = rc.right;
        vertex[1].y = rc.bottom;
        vertex[1].Red = static_cast<COLOR16>(GetRValue(bottomColor) << 8);
        vertex[1].Green = static_cast<COLOR16>(GetGValue(bottomColor) << 8);
        vertex[1].Blue = static_cast<COLOR16>(GetBValue(bottomColor) << 8);
        vertex[1].Alpha = 0;

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
            titleFont_(nullptr),
            textFont_(nullptr),
            selectedTextFont_(nullptr),
            smallFont_(nullptr) {}

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
        460,
        200,
        nullptr,
        nullptr,
        g_moduleHandle,
        this);

    EnsureFonts();
    return hwnd_ != nullptr;
}

bool CandidateWindow::EnsureFonts() {
    if (titleFont_ != nullptr && textFont_ != nullptr && selectedTextFont_ != nullptr && smallFont_ != nullptr) {
        return true;
    }

    if (titleFont_ == nullptr) {
        titleFont_ = CreateFontW(-16, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
            OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Microsoft YaHei UI");
    }
    if (textFont_ == nullptr) {
        textFont_ = CreateFontW(-17, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
            OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Microsoft YaHei UI");
    }
    if (selectedTextFont_ == nullptr) {
        selectedTextFont_ = CreateFontW(-18, 0, 0, 0, FW_SEMIBOLD, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
            OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Microsoft YaHei UI");
    }
    if (smallFont_ == nullptr) {
        smallFont_ = CreateFontW(-14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE, DEFAULT_CHARSET,
            OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, CLEARTYPE_QUALITY, DEFAULT_PITCH | FF_SWISS, L"Microsoft YaHei UI");
    }

    return titleFont_ != nullptr && textFont_ != nullptr && selectedTextFont_ != nullptr && smallFont_ != nullptr;
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
}

void CandidateWindow::Destroy() {
    if (hwnd_ != nullptr) {
        DestroyWindow(hwnd_);
        hwnd_ = nullptr;
    }
    ReleaseFonts();
    hasLastWindowRect_ = false;
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

    bool candidatesChanged = false;
    if (!cheapStateChanged) {
        candidatesChanged = !AreDisplayCandidatesEqual(candidates_, candidates);
    }

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

    int x = 0;
    int y = 0;
    POINT anchorPoint = {0, 0};
    int anchorGap = 14;
    bool usedValidAnchor = false;
    if (anchorScreenPos != nullptr) {
        anchorPoint = *anchorScreenPos;
        x = anchorPoint.x + 10;
        y = anchorPoint.y + anchorGap;
        usedValidAnchor = true;
    }

    if (!usedValidAnchor) {
        POINT cursor = {};
        if (GetCursorPos(&cursor)) {
            anchorPoint = cursor;
            anchorGap = 20;
            x = anchorPoint.x + 12;
            y = anchorPoint.y + anchorGap;
            usedValidAnchor = true;
        }
    }

    if (!usedValidAnchor) {
        x = 120;
        y = 120;
    }

    EnsureFonts();
    int width = hasLastWindowRect_ ? lastWidth_ : 300;
    if (contentChanged || !hasLastWindowRect_) {
        width = 300;
        HDC measureDc = GetDC(hwnd_);
        if (measureDc != nullptr) {
            HFONT measureTextFont = textFont_ != nullptr ? textFont_ : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
            HFONT measureSmallFont = smallFont_ != nullptr ? smallFont_ : measureTextFont;
            if (measureTextFont != nullptr && measureSmallFont != nullptr) {
                HGDIOBJ oldFont = SelectObject(measureDc, measureTextFont);
                SIZE textSize = {};
                SIZE codeSize = {};

                const size_t measureRows = std::min(candidates_.size(), kMaxVisibleRows);
                for (size_t i = 0; i < measureRows; ++i) {
                    const std::wstring& candidateCode = candidates_[i].code;
                    const wchar_t* codeTextPtr = candidateCode.empty() ? L"-" : candidateCode.c_str();
                    const int codeTextLen = candidateCode.empty() ? 1 : static_cast<int>(candidateCode.size());
                    GetTextExtentPoint32W(measureDc, candidates_[i].text.c_str(), static_cast<int>(candidates_[i].text.size()), &textSize);
                    SelectObject(measureDc, measureSmallFont);
                    GetTextExtentPoint32W(measureDc, codeTextPtr, codeTextLen, &codeSize);
                    SelectObject(measureDc, measureTextFont);

                    const int codeBlockWidth = std::max(96, std::min(static_cast<int>(codeSize.cx) + 22, 280));
                    const int neededWidth = 36 + static_cast<int>(textSize.cx) + 16 + codeBlockWidth + 22;
                    width = std::max(width, neededWidth);
                }

                SelectObject(measureDc, oldFont);
            }
            ReleaseDC(hwnd_, measureDc);
        }
        width = std::max(300, std::min(width, 960));
    }

    const int headerHeight = 30;
    const int codeHeight = 30;
    const int rowHeight = 32;
    const int footerHeight = 44;
    int height = hasLastWindowRect_ ? lastHeight_ : (headerHeight + codeHeight + footerHeight + rowHeight + 14);
    if (contentChanged || !hasLastWindowRect_) {
        const int rowCount = std::max(1, static_cast<int>(std::min(candidates_.size(), kMaxVisibleRows)));
        const int cappedRows = std::max(1, rowCount);
        height = headerHeight + codeHeight + footerHeight + rowHeight * cappedRows + 14;
    }

    RECT workArea = {};
    POINT workAreaPoint = usedValidAnchor ? anchorPoint : POINT{x, y};
    if (GetWorkAreaForPoint(workAreaPoint, workArea)) {
        if (!usedValidAnchor && x == 120 && y == 120) {
            x = static_cast<int>(workArea.right) - width - 24;
            y = static_cast<int>(workArea.bottom) - height - 36;
        }

        if (usedValidAnchor) {
            const int minY = static_cast<int>(workArea.top);
            const int maxY = std::max(minY, static_cast<int>(workArea.bottom) - height);
            const int belowY = anchorPoint.y + anchorGap;
            const int aboveY = anchorPoint.y - height - 12;
            const int clampedBelowY = std::max(minY, std::min(belowY, maxY));
            const int clampedAboveY = std::max(minY, std::min(aboveY, maxY));

            const int anchorBandTop = anchorPoint.y - 16;
            const int anchorBandBottom = anchorPoint.y + 26;
            const auto overlapWithAnchorBand = [&](int topY) {
                const int bottomY = topY + height;
                const int overlapTop = std::max(topY, anchorBandTop);
                const int overlapBottom = std::min(bottomY, anchorBandBottom);
                return std::max(0, overlapBottom - overlapTop);
            };

            const int belowOverlap = overlapWithAnchorBand(clampedBelowY);
            const int aboveOverlap = overlapWithAnchorBand(clampedAboveY);

            if (aboveOverlap < belowOverlap) {
                y = clampedAboveY;
            } else if (belowOverlap < aboveOverlap) {
                y = clampedBelowY;
            } else {
                const int belowAdjust = std::abs(clampedBelowY - belowY);
                const int aboveAdjust = std::abs(clampedAboveY - aboveY);
                y = (aboveAdjust <= belowAdjust) ? clampedAboveY : clampedBelowY;
            }
        }

        const int minX = static_cast<int>(workArea.left);
        const int maxX = std::max(minX, static_cast<int>(workArea.right) - width);
        const int minY = static_cast<int>(workArea.top);
        const int maxY = std::max(minY, static_cast<int>(workArea.bottom) - height);

        if (x < minX) {
            x = minX;
        }
        if (x > maxX) {
            x = maxX;
        }
        if (y < minY) {
            y = minY;
        }
        if (y > maxY) {
            y = maxY;
        }
    }

    int smoothedX = x;
    int smoothedY = y;
    int smoothedWidth = width;
    int smoothedHeight = height;
    const bool preferCaretFollow = usedValidAnchor;

    if (!hasLastWindowRect_) {
        hasLastWindowRect_ = true;
        lastX_ = x;
        lastY_ = y;
        lastWidth_ = width;
        lastHeight_ = height;
    } else if (preferCaretFollow) {
        smoothedX = x;
        smoothedY = y;
        smoothedWidth = ApproachIntElastic(lastWidth_, width, 8, 0.56, 0.30, 6);
        smoothedHeight = ApproachIntElastic(lastHeight_, height, 8, 0.56, 0.30, 6);

        POINT clampPoint = usedValidAnchor ? anchorPoint : POINT{x, y};
        if (GetWorkAreaForPoint(clampPoint, workArea)) {
            const int maxX = std::max(static_cast<int>(workArea.left), static_cast<int>(workArea.right) - smoothedWidth);
            const int maxY = std::max(static_cast<int>(workArea.top), static_cast<int>(workArea.bottom) - smoothedHeight);
            smoothedX = std::max(static_cast<int>(workArea.left), std::min(smoothedX, maxX));
            smoothedY = std::max(static_cast<int>(workArea.top), std::min(smoothedY, maxY));
        }

        lastX_ = smoothedX;
        lastY_ = smoothedY;
        lastWidth_ = smoothedWidth;
        lastHeight_ = smoothedHeight;
    } else {
        const int moveDelta = std::abs(x - lastX_) + std::abs(y - lastY_);
        const bool farJump = moveDelta > 240;

        const int moveMinStep = farJump ? 30 : 10;
        const double moveRatio = farJump ? 0.62 : 0.36;
        const int sizeMinStep = 8;
        const double sizeGrowRatio = 0.56;
        const double sizeShrinkRatio = 0.30;
        const int moveSnapDistance = 5;
        const int sizeSnapDistance = 6;

        smoothedX = farJump
            ? ApproachInt(lastX_, x, moveMinStep, moveRatio)
            : ApproachIntSpring(lastX_, x, moveMinStep, moveRatio, moveSnapDistance, 18, 6);
        smoothedY = farJump
            ? ApproachInt(lastY_, y, moveMinStep, moveRatio)
            : ApproachIntSpring(lastY_, y, moveMinStep, moveRatio, moveSnapDistance, 18, 6);
        smoothedWidth = ApproachIntElastic(lastWidth_, width, sizeMinStep, sizeGrowRatio, sizeShrinkRatio, sizeSnapDistance);
        smoothedHeight = ApproachIntElastic(lastHeight_, height, sizeMinStep, sizeGrowRatio, sizeShrinkRatio, sizeSnapDistance);

        if (std::abs(width - smoothedWidth) <= 20) {
            smoothedWidth = ApproachIntSpring(smoothedWidth, width, 3, 0.40, sizeSnapDistance, 14, 5);
        }
        if (std::abs(height - smoothedHeight) <= 20) {
            smoothedHeight = ApproachIntSpring(smoothedHeight, height, 3, 0.40, sizeSnapDistance, 14, 5);
        }

        POINT clampPoint = usedValidAnchor ? anchorPoint : POINT{x, y};
        if (GetWorkAreaForPoint(clampPoint, workArea)) {
            const int maxX = std::max(static_cast<int>(workArea.left), static_cast<int>(workArea.right) - smoothedWidth);
            const int maxY = std::max(static_cast<int>(workArea.top), static_cast<int>(workArea.bottom) - smoothedHeight);
            smoothedX = std::max(static_cast<int>(workArea.left), std::min(smoothedX, maxX));
            smoothedY = std::max(static_cast<int>(workArea.top), std::min(smoothedY, maxY));
        }

        lastX_ = smoothedX;
        lastY_ = smoothedY;
        lastWidth_ = smoothedWidth;
        lastHeight_ = smoothedHeight;
    }

    bool hadWindowRect = false;
    RECT currentRect = {};
    if (GetWindowRect(hwnd_, &currentRect)) {
        hadWindowRect = true;
    }
    const bool wasVisible = IsWindowVisible(hwnd_) == TRUE;

    const bool movedOrResized =
        !hadWindowRect ||
        currentRect.left != smoothedX ||
        currentRect.top != smoothedY ||
        (currentRect.right - currentRect.left) != smoothedWidth ||
        (currentRect.bottom - currentRect.top) != smoothedHeight;
    const bool sizeChanged =
        !hadWindowRect ||
        (currentRect.right - currentRect.left) != smoothedWidth ||
        (currentRect.bottom - currentRect.top) != smoothedHeight;

    if (movedOrResized || !wasVisible) {
        SetWindowPos(hwnd_, HWND_TOP, smoothedX, smoothedY, smoothedWidth, smoothedHeight, SWP_SHOWWINDOW | SWP_NOACTIVATE);
    }

    if (contentChanged || sizeChanged || !wasVisible) {
        InvalidateRect(hwnd_, nullptr, TRUE);
    }
}

void CandidateWindow::Hide() {
    if (hwnd_ != nullptr) {
        ShowWindow(hwnd_, SW_HIDE);
    }
}

LRESULT CALLBACK CandidateWindow::WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam) {
    if (msg == WM_NCCREATE) {
        auto* cs = reinterpret_cast<CREATESTRUCTW*>(lParam);
        auto* self = static_cast<CandidateWindow*>(cs->lpCreateParams);
        SetWindowLongPtrW(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(self));
        return DefWindowProcW(hwnd, msg, wParam, lParam);
    }

    auto* self = reinterpret_cast<CandidateWindow*>(GetWindowLongPtrW(hwnd, GWLP_USERDATA));

    switch (msg) {
    case WM_PAINT:
        if (self != nullptr) {
            self->OnPaint();
            return 0;
        }
        break;
    case WM_ERASEBKGND:
        return 1;
    default:
        break;
    }

    return DefWindowProcW(hwnd, msg, wParam, lParam);
}

void CandidateWindow::OnPaint() {
    EnsureFonts();
    PAINTSTRUCT ps = {};
    HDC hdc = BeginPaint(hwnd_, &ps);

    RECT rc = {};
    GetClientRect(hwnd_, &rc);

    const COLORREF bg = RGB(248, 245, 238);
    const COLORREF panelBg = RGB(255, 253, 249);
    const COLORREF titleBg = RGB(246, 227, 198);
    const COLORREF border = RGB(201, 177, 144);
    const COLORREF textMain = RGB(44, 39, 34);
    const COLORREF textMeta = RGB(118, 103, 90);
    const COLORREF indexBg = RGB(255, 235, 208);
    const COLORREF indexFg = RGB(156, 88, 23);
    const COLORREF indexBorder = RGB(227, 177, 114);
    const COLORREF activeRowBg = RGB(255, 236, 211);
    const COLORREF activeRowBorder = RGB(215, 141, 64);
    const COLORREF activeBar = RGB(226, 118, 26);
    const COLORREF cardBg = RGB(255, 253, 248);
    const COLORREF cardBorder = RGB(237, 220, 197);
    const COLORREF cardBorderNear = RGB(229, 197, 158);
    const COLORREF cardTextFar = RGB(108, 92, 74);

    size_t userHitsTop3 = 0;
    size_t learnedHitsTop3 = 0;
    const size_t topEvalCount = std::min(static_cast<size_t>(3), candidates_.size());
    for (size_t i = 0; i < topEvalCount; ++i) {
        if (candidates_[i].boostedUser) {
            ++userHitsTop3;
        }
        else if (candidates_[i].boostedLearned) {
            ++learnedHitsTop3;
        }
    }

    const bool topExact = !candidates_.empty() && !code_.empty() && candidates_[0].code == code_;
    const bool selectedValid = !candidates_.empty() && selectedIndex_ < candidates_.size();
    const bool selectedExact = selectedValid && !code_.empty() && candidates_[selectedIndex_].code == code_;
    const bool selectedUser = selectedValid && candidates_[selectedIndex_].boostedUser;
    const bool selectedLearned = selectedValid && candidates_[selectedIndex_].boostedLearned;
    const bool selectedContext = selectedValid && candidates_[selectedIndex_].boostedContext;
    int qualityLevel = 0;  // 0=LOW, 1=MED, 2=HIGH
    std::wstring qualityReason = L"base/jichu";
    std::wstring qualityReasonBadge = L"BS/JC";
    std::wstring qualityText = L"Q: LOW";
    std::wstring qualityShort = L"L";
    COLORREF qualityBg = RGB(239, 242, 248);
    COLORREF qualityBorder = RGB(201, 210, 226);
    COLORREF qualityFg = RGB(106, 118, 142);
    if ((topExact && userHitsTop3 > 0) || (selectedExact && selectedUser)) {
        qualityLevel = 2;
        qualityReason = selectedExact && selectedUser ? L"user/yonghu" : L"exact/jingque";
        qualityReasonBadge = selectedExact && selectedUser ? L"US/YH" : L"EX/JQ";
        qualityText = L"Q: HIGH";
        qualityShort = L"H";
        qualityBg = RGB(215, 233, 255);
        qualityBorder = RGB(145, 182, 236);
        qualityFg = RGB(35, 79, 154);
    }
    else if (topExact || learnedHitsTop3 > 0 || selectedExact || selectedLearned || selectedContext) {
        qualityLevel = 1;
        qualityReason = selectedContext ? L"assoc/lianxiang" : (selectedLearned ? L"learn/xuexi" : L"exact/jingque");
        qualityReasonBadge = selectedContext ? L"AS/LX" : (selectedLearned ? L"LR/XX" : L"EX/JQ");
        qualityText = L"Q: MED";
        qualityShort = L"M";
        qualityBg = RGB(226, 240, 255);
        qualityBorder = RGB(168, 196, 236);
        qualityFg = RGB(52, 91, 161);
    }

    if (selectedValid && selectedIndex_ > 2 && !selectedExact && !selectedUser) {
        if (qualityLevel == 2) {
            qualityLevel = 1;
            qualityReason = L"gap/chaju";
            qualityReasonBadge = L"GP/CJ";
            qualityText = L"Q: MED";
            qualityShort = L"M";
            qualityBg = RGB(226, 240, 255);
            qualityBorder = RGB(168, 196, 236);
            qualityFg = RGB(52, 91, 161);
        }
        else if (qualityLevel == 1) {
            qualityLevel = 0;
            qualityReason = L"gap/chaju";
            qualityReasonBadge = L"GP/CJ";
            qualityText = L"Q: LOW";
            qualityShort = L"L";
            qualityBg = RGB(239, 242, 248);
            qualityBorder = RGB(201, 210, 226);
            qualityFg = RGB(106, 118, 142);
        }
    }

    const ULONGLONG nowTick = GetTickCount64();
    const bool hadPreviousQuality = qualityLevelTick_ != 0;
    const int prevDisplayedQualityLevel = displayedQualityLevel_;
    bool qualityChanged = false;
    if (qualityLevelTick_ == 0) {
        displayedQualityLevel_ = qualityLevel;
        qualityLevelTick_ = nowTick;
        qualityChanged = true;
    }
    else if (qualityLevel > displayedQualityLevel_) {
        displayedQualityLevel_ = qualityLevel;
        qualityLevelTick_ = nowTick;
        qualityChanged = true;
    }
    else if (qualityLevel < displayedQualityLevel_) {
        if ((nowTick - qualityLevelTick_) >= 180ULL) {
            displayedQualityLevel_ = qualityLevel;
            qualityLevelTick_ = nowTick;
            qualityChanged = true;
        }
    }

    int effectiveQualityLevel = displayedQualityLevel_;
    if (effectiveQualityLevel < 0) {
        effectiveQualityLevel = 0;
    }
    if (effectiveQualityLevel > 2) {
        effectiveQualityLevel = 2;
    }

    if (effectiveQualityLevel != qualityLevel) {
        qualityReason += L"-hold";
        qualityReasonBadge += L"+H";
    }

    const bool qualityUpgraded = qualityChanged && hadPreviousQuality && displayedQualityLevel_ > prevDisplayedQualityLevel;
    if (qualityUpgraded) {
        const int levelDelta = displayedQualityLevel_ - prevDisplayedQualityLevel;
        ULONGLONG pulseMs = 160ULL;
        if (levelDelta >= 2) {
            pulseMs = 240ULL;
        }
        else if (levelDelta == 1) {
            pulseMs = 170ULL;
        }
        qualityPulseUntilTick_ = nowTick + pulseMs;
    }

    if (effectiveQualityLevel == 2) {
        qualityText = L"Q: HIGH";
        qualityShort = L"H";
        qualityBg = RGB(215, 233, 255);
        qualityBorder = RGB(145, 182, 236);
        qualityFg = RGB(35, 79, 154);
    }
    else if (effectiveQualityLevel == 1) {
        qualityText = L"Q: MED";
        qualityShort = L"M";
        qualityBg = RGB(226, 240, 255);
        qualityBorder = RGB(168, 196, 236);
        qualityFg = RGB(52, 91, 161);
    }
    else {
        qualityText = L"Q: LOW";
        qualityShort = L"L";
        qualityBg = RGB(239, 242, 248);
        qualityBorder = RGB(201, 210, 226);
        qualityFg = RGB(106, 118, 142);
    }

    FillSolidRect(hdc, rc, bg);

    RECT panelRc = rc;
    panelRc.left += 1;
    panelRc.top += 1;
    panelRc.right -= 1;
    panelRc.bottom -= 1;

    FillSolidRect(hdc, panelRc, panelBg);
    FillVerticalGradient(hdc, panelRc, RGB(255, 254, 252), RGB(253, 248, 241));

    HPEN borderPen = CreatePen(PS_SOLID, 1, border);
    HGDIOBJ oldPen = SelectObject(hdc, borderPen);
    HGDIOBJ oldBrush = SelectObject(hdc, GetStockObject(HOLLOW_BRUSH));
    RoundRect(hdc, panelRc.left, panelRc.top, panelRc.right, panelRc.bottom, 12, 12);
    SelectObject(hdc, oldBrush);
    SelectObject(hdc, oldPen);
    DeleteObject(borderPen);

    RECT titleRc = panelRc;
    titleRc.bottom = titleRc.top + 30;
    FillSolidRect(hdc, titleRc, titleBg);
    FillVerticalGradient(hdc, titleRc, RGB(253, 239, 217), RGB(243, 218, 180));

    SetBkMode(hdc, TRANSPARENT);

    HFONT titleFont = titleFont_ != nullptr ? titleFont_ : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    HFONT textFont = textFont_ != nullptr ? textFont_ : static_cast<HFONT>(GetStockObject(DEFAULT_GUI_FONT));
    HFONT selectedTextFont = selectedTextFont_ != nullptr ? selectedTextFont_ : textFont;
    HFONT smallFont = smallFont_ != nullptr ? smallFont_ : textFont;

    HGDIOBJ oldFont = SelectObject(hdc, titleFont);
    SetTextColor(hdc, textMain);

    const std::wstring title = L"Yun";

    RECT textRc = titleRc;
    textRc.left += 12;
    textRc.right -= 250;
    DrawTextW(hdc, title.c_str(), static_cast<int>(title.size()), &textRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

    wchar_t pageTextBuffer[96] = {};
    swprintf_s(
        pageTextBuffer,
        L"Page %zu/%zu  Total %zu",
        totalPages_ == 0 ? static_cast<size_t>(0) : (pageIndex_ + 1),
        totalPages_,
        totalCandidateCount_);
    const bool canPrevPage = (pageIndex_ > 0);
    const bool canNextPage = (totalPages_ > 0 && pageIndex_ + 1 < totalPages_);
    RECT pageRc = titleRc;
    pageRc.right -= 12;
    pageRc.left = pageRc.right - 178;
    pageRc.top += 5;
    pageRc.bottom -= 5;
    DrawRoundedRect(hdc, pageRc, 12, 12, RGB(255, 223, 179), border);
    SetTextColor(hdc, RGB(138, 78, 24));
    DrawTextW(hdc, pageTextBuffer, -1, &pageRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

    // Removed top-right quality dot and pulse indicator per UI request.

    RECT pageNavRc = titleRc;
    pageNavRc.right = pageRc.left - 8;
    pageNavRc.left = pageNavRc.right - 86;
    pageNavRc.top += 6;
    pageNavRc.bottom -= 6;
    SetTextColor(hdc, RGB(150, 121, 91));
    const std::wstring pageHintText;
    DrawTextW(hdc, pageHintText.c_str(), static_cast<int>(pageHintText.size()), &pageNavRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

    if (totalPages_ > 1) {
        RECT pagerStripRc = titleRc;
        pagerStripRc.left += 14;
        pagerStripRc.right = pageNavRc.left - 8;
        pagerStripRc.top = titleRc.bottom - 6;
        pagerStripRc.bottom = pagerStripRc.top + 3;

        FillSolidRect(hdc, pagerStripRc, RGB(206, 221, 245));

        RECT pagerFillRc = pagerStripRc;
        const int stripWidth = std::max<int>(1, static_cast<int>(pagerStripRc.right - pagerStripRc.left));
        const int fillWidth = static_cast<int>((static_cast<double>(pageIndex_ + 1) / static_cast<double>(totalPages_)) * stripWidth);
        pagerFillRc.right = pagerFillRc.left + std::max(10, fillWidth);
        FillSolidRect(hdc, pagerFillRc, RGB(84, 128, 208));
    }

    SelectObject(hdc, textFont);
    RECT codeRc = panelRc;
    codeRc.top = titleRc.bottom + 4;
    codeRc.left += 12;
    codeRc.right -= 12;
    codeRc.bottom = codeRc.top + 26;

    FillSolidRect(hdc, codeRc, RGB(255, 248, 236));
    DrawRectBorder(hdc, codeRc, border);

    SetTextColor(hdc, textMain);
    std::wstring codeLine = L"Code: " + code_;
    RECT codeTextRc = codeRc;
    codeTextRc.left += 8;
    DrawTextW(hdc, codeLine.c_str(), static_cast<int>(codeLine.size()), &codeTextRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

    if (!code_.empty()) {
        SIZE prefixSize = {};
        const std::wstring prefix = L"Code: ";
        GetTextExtentPoint32W(hdc, prefix.c_str(), static_cast<int>(prefix.size()), &prefixSize);

        RECT codeHighlightRc = codeRc;
        codeHighlightRc.left = codeTextRc.left + prefixSize.cx - 2;
        codeHighlightRc.right = std::min(codeHighlightRc.left + static_cast<int>(code_.size()) * 13 + 12, codeRc.right - 6);
        codeHighlightRc.top += 4;
        codeHighlightRc.bottom -= 4;
        FillSolidRect(hdc, codeHighlightRc, RGB(255, 229, 190));

        SetTextColor(hdc, RGB(164, 86, 16));
        RECT codeOnlyRc = codeHighlightRc;
        codeOnlyRc.left += 6;
        DrawTextW(hdc, code_.c_str(), static_cast<int>(code_.size()), &codeOnlyRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE);
    }

    const int rowLeft = panelRc.left + 12;
    const int rowRight = panelRc.right - 12;
    const int rowWidth = rowRight - rowLeft;
    int y = codeRc.bottom + 6;
    if (candidates_.empty()) {
        SelectObject(hdc, textFont);
        SetTextColor(hdc, textMeta);
        RECT emptyRc = panelRc;
        emptyRc.top = y;
        emptyRc.left += 12;
        emptyRc.right -= 12;
        emptyRc.bottom = emptyRc.top + 24;
        const std::wstring empty = L"No candidates for current code";
        DrawTextW(hdc, empty.c_str(), static_cast<int>(empty.size()), &emptyRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

        RECT emptyHintRc = emptyRc;
        emptyHintRc.top = emptyRc.bottom + 4;
        emptyHintRc.bottom = emptyHintRc.top + 20;
        SelectObject(hdc, smallFont);
        SetTextColor(hdc, RGB(96, 114, 146));
        const std::wstring emptyHint = L"Space/Enter raw code | Digit raw fallback | Backspace edit | Esc cancel";
        DrawTextW(hdc, emptyHint.c_str(), static_cast<int>(emptyHint.size()), &emptyHintRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);
        y = emptyHintRc.bottom + 6;
    } else {
        const size_t maxRows = std::min(candidates_.size(), kMaxVisibleRows);
        for (size_t i = 0; i < maxRows; ++i) {
            RECT rowRc = {rowLeft, y, rowLeft + rowWidth, y + 30};
            const bool rowSelected = (i == selectedIndex_);
            const int selectionDistance = std::abs(static_cast<int>(i) - static_cast<int>(selectedIndex_));

            RECT cardRc = rowRc;
            cardRc.left += 2;
            cardRc.right -= 2;

            if (rowSelected) {
                RECT shadowRc = cardRc;
                shadowRc.left += 2;
                shadowRc.right -= 2;
                shadowRc.top += 2;
                shadowRc.bottom += 2;
                FillSolidRect(hdc, shadowRc, RGB(246, 225, 199));
            }

            if (rowSelected) {
                DrawRoundedRect(hdc, cardRc, 8, 8, activeRowBg, activeRowBorder);
                FillVerticalGradient(hdc, cardRc, RGB(255, 242, 223), RGB(255, 229, 198));
            }
            else if (selectionDistance <= 1) {
                const COLORREF borderTone = cardBorderNear;
                DrawRoundedRect(hdc, cardRc, 8, 8, cardBg, borderTone);
                FillVerticalGradient(hdc, cardRc, RGB(255, 254, 250), RGB(252, 245, 235));
            }
            else {
                const COLORREF borderTone = cardBorder;
                DrawRoundedRect(hdc, cardRc, 8, 8, cardBg, borderTone);
                FillVerticalGradient(hdc, cardRc, RGB(255, 253, 249), RGB(251, 244, 235));
            }

            if (rowSelected) {
                RECT activeBarRc = cardRc;
                activeBarRc.left += 1;
                activeBarRc.right = activeBarRc.left + 4;
                activeBarRc.top += 2;
                activeBarRc.bottom -= 2;
                FillSolidRect(hdc, activeBarRc, activeBar);
            }

            const wchar_t* indexText = i < (sizeof(kIndexLabels) / sizeof(kIndexLabels[0])) ? kIndexLabels[i] : L"?";
            RECT idxRc = cardRc;
            idxRc.right = idxRc.left + 26;
            idxRc.top += 3;
            idxRc.bottom -= 3;

            DrawRoundedRect(hdc, idxRc, 8, 8, indexBg, indexBorder);

            RECT idxHighlightRc = idxRc;
            idxHighlightRc.left += 2;
            idxHighlightRc.right -= 2;
            idxHighlightRc.top += 2;
            idxHighlightRc.bottom = idxHighlightRc.top + 2;
            FillSolidRect(hdc, idxHighlightRc, RGB(246, 250, 255));

            SetTextColor(hdc, indexFg);
            SelectObject(hdc, smallFont);
            DrawTextW(hdc, indexText, -1, &idxRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

            RECT textRc = cardRc;
            textRc.left += 38;
            const std::wstring& candidateCode = candidates_[i].code;
            const wchar_t* codeTextPtr = candidateCode.empty() ? L"-" : candidateCode.c_str();
            const int codeTextLen = candidateCode.empty() ? 1 : static_cast<int>(candidateCode.size());

            SelectObject(hdc, smallFont);
            SIZE codeSize = {};
            GetTextExtentPoint32W(hdc, codeTextPtr, codeTextLen, &codeSize);
            int codeTagWidth = std::max(96, std::min(static_cast<int>(codeSize.cx) + 22, 220));

            textRc.right = cardRc.right - codeTagWidth - 16;
            if (rowSelected) {
                SetTextColor(hdc, RGB(130, 69, 14));
                SelectObject(hdc, selectedTextFont != nullptr ? selectedTextFont : textFont);
            }
            else if (selectionDistance <= 1) {
                SetTextColor(hdc, textMain);
                SelectObject(hdc, textFont);
            }
            else {
                SetTextColor(hdc, cardTextFar);
                SelectObject(hdc, textFont);
            }
            DrawTextW(hdc, candidates_[i].text.c_str(), static_cast<int>(candidates_[i].text.size()), &textRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

            RECT codeTagRc = cardRc;
            codeTagRc.left = cardRc.right - codeTagWidth - 6;
            codeTagRc.right = cardRc.right - 6;
            codeTagRc.top += 5;
            codeTagRc.bottom -= 5;

            RECT dividerRc = cardRc;
            dividerRc.left = codeTagRc.left - 8;
            dividerRc.right = dividerRc.left + 1;
            dividerRc.top += 5;
            dividerRc.bottom -= 5;
            FillSolidRect(hdc, dividerRc, rowSelected ? RGB(225, 164, 94) : RGB(232, 212, 185));

            const COLORREF codeTagBg = rowSelected ? RGB(207, 120, 28) : RGB(255, 240, 216);
            const COLORREF codeTagBorder = rowSelected ? RGB(170, 96, 21) : RGB(234, 201, 160);
            const COLORREF codeTagText = rowSelected ? RGB(255, 250, 240) : textMeta;

            DrawRoundedRect(hdc, codeTagRc, 12, 12, codeTagBg, codeTagBorder);

            SetTextColor(hdc, codeTagText);
            SelectObject(hdc, smallFont);
            RECT codeTextRc = codeTagRc;
            codeTextRc.left += 10;
            codeTextRc.right -= 8;
            DrawTextW(hdc, codeTextPtr, codeTextLen, &codeTextRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

            y += 32;
        }
    }

    SelectObject(hdc, smallFont);
    SetTextColor(hdc, textMeta);

    RECT infoRc = panelRc;
    infoRc.left += 12;
    infoRc.right -= 12;
    infoRc.bottom -= 28;
    infoRc.top = infoRc.bottom - 16;

    std::wstring infoLine = L"Focus: none";
    if (!candidates_.empty()) {
        const size_t safeIdx = std::min(selectedIndex_, candidates_.size() - 1);
        const std::wstring selectedText = candidates_[safeIdx].text.empty() ? L"-" : candidates_[safeIdx].text;
        const std::wstring selectedCode = candidates_[safeIdx].code.empty() ? L"-" : candidates_[safeIdx].code;
        std::wstring source = L"BASE";
        if (candidates_[safeIdx].boostedUser) {
            source = L"USER";
        }
        else if (candidates_[safeIdx].boostedContext) {
            source = L"ASSOC";
        }

        const size_t consumedLength = std::min(candidates_[safeIdx].consumedLength, code_.size());
        const bool isFirstLevelShortCode =
            consumedLength == 1 &&
            code_.size() == 1 &&
            !candidates_[safeIdx].code.empty() &&
            candidates_[safeIdx].code.size() == 1;
        const std::wstring matchMode =
            isFirstLevelShortCode ? L"JM1" :
            (consumedLength < code_.size() ? L"CONT" : (candidates_[safeIdx].boostedContext ? L"ASSOC" : L"FULL"));
        infoLine = L"Focus #" + std::to_wstring(std::min(selectedAbsoluteIndex_ + 1, totalCandidateCount_ == 0 ? static_cast<size_t>(1) : totalCandidateCount_)) +
            L"/" + std::to_wstring(totalCandidateCount_) + L" (key " + std::to_wstring(safeIdx + 1) + L", " + source + L", " + matchMode + L"): " + selectedText +
            L"  [" + selectedCode + L"]";
    }

    const std::wstring learnHint = L"Learn: freq on | assoc on | continue on | Pin Ctrl+1-9 | Delete Ctrl+Del";

    FillSolidRect(hdc, infoRc, RGB(247, 251, 255));
    DrawRectBorder(hdc, infoRc, RGB(221, 232, 248));

    RECT infoTextRc = infoRc;
    infoTextRc.left += 8;
    infoTextRc.right -= 380;
    DrawTextW(hdc, infoLine.c_str(), static_cast<int>(infoLine.size()), &infoTextRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

    RECT reasonRc = infoRc;
    reasonRc.left = infoTextRc.right + 8;
    reasonRc.right = reasonRc.left + 88;
    reasonRc.top += 2;
    reasonRc.bottom -= 2;

    COLORREF reasonBg = RGB(232, 242, 255);
    COLORREF reasonBorder = RGB(171, 198, 239);
    COLORREF reasonText = RGB(49, 92, 161);
    if (qualityReasonBadge.find(L"US/") == 0) {
        reasonBg = RGB(221, 236, 255);
        reasonBorder = RGB(145, 185, 236);
        reasonText = RGB(37, 80, 153);
    }
    else if (qualityReasonBadge.find(L"LR/") == 0) {
        reasonBg = RGB(230, 242, 255);
        reasonBorder = RGB(164, 197, 236);
        reasonText = RGB(54, 92, 159);
    }
    else if (qualityReasonBadge.find(L"AS/") == 0) {
        reasonBg = RGB(232, 245, 243);
        reasonBorder = RGB(166, 210, 202);
        reasonText = RGB(52, 121, 109);
    }
    else if (qualityReasonBadge.find(L"GP/") == 0) {
        reasonBg = RGB(245, 237, 232);
        reasonBorder = RGB(212, 177, 152);
        reasonText = RGB(137, 93, 68);
    }
    else if (qualityReasonBadge.find(L"EX/") == 0) {
        reasonBg = RGB(225, 239, 255);
        reasonBorder = RGB(154, 190, 236);
        reasonText = RGB(44, 87, 158);
    }

    DrawRoundedRect(hdc, reasonRc, 10, 10, reasonBg, reasonBorder);
    SetTextColor(hdc, reasonText);
    DrawTextW(hdc, qualityReasonBadge.c_str(), static_cast<int>(qualityReasonBadge.size()), &reasonRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

    RECT qualityRc = infoRc;
    qualityRc.left = reasonRc.right + 8;
    qualityRc.right = qualityRc.left + 76;
    qualityRc.top += 2;
    qualityRc.bottom -= 2;
    DrawRoundedRect(hdc, qualityRc, 10, 10, qualityBg, qualityBorder);
    SetTextColor(hdc, qualityFg);
    DrawTextW(hdc, qualityText.c_str(), static_cast<int>(qualityText.size()), &qualityRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

    RECT learnHintRc = infoRc;
    learnHintRc.left = qualityRc.right + 8;
    learnHintRc.right -= 8;
    SetTextColor(hdc, RGB(89, 109, 146));
    DrawTextW(hdc, learnHint.c_str(), static_cast<int>(learnHint.size()), &learnHintRc, DT_RIGHT | DT_VCENTER | DT_SINGLELINE | DT_END_ELLIPSIS);

    RECT footerRc = panelRc;
    footerRc.left += 12;
    footerRc.right -= 12;
    footerRc.bottom -= 6;
    footerRc.top = footerRc.bottom - 16;

    RECT boundaryRc = footerRc;
    boundaryRc.bottom = footerRc.top - 2;
    boundaryRc.top = boundaryRc.bottom - 14;
    FillSolidRect(hdc, boundaryRc, RGB(246, 250, 255));
    DrawRectBorder(hdc, boundaryRc, RGB(221, 232, 248));

    std::wstring boundaryText = L"MID PAGE";
    COLORREF boundaryFg = RGB(108, 127, 161);
    if (!canPrevPage) {
        boundaryText = L"FIRST PAGE";
        boundaryFg = RGB(56, 99, 175);
    }
    if (!canNextPage) {
        boundaryText = canPrevPage ? L"LAST PAGE" : L"FIRST/LAST PAGE";
        boundaryFg = RGB(56, 99, 175);
    }
    SetTextColor(hdc, boundaryFg);
    DrawTextW(hdc, boundaryText.c_str(), static_cast<int>(boundaryText.size()), &boundaryRc, DT_CENTER | DT_VCENTER | DT_SINGLELINE);

    const bool hasCandidates = !candidates_.empty();
    auto drawFooterPart = [&](const std::wstring& text, COLORREF color, int& cursorX) {
        RECT partRc = footerRc;
        partRc.left = cursorX;
        partRc.right = footerRc.right;
        SetTextColor(hdc, color);
        DrawTextW(hdc, text.c_str(), static_cast<int>(text.size()), &partRc, DT_LEFT | DT_VCENTER | DT_SINGLELINE);

        SIZE partSize = {};
        GetTextExtentPoint32W(hdc, text.c_str(), static_cast<int>(text.size()), &partSize);
        cursorX += static_cast<int>(partSize.cx);
    };

    int footerX = footerRc.left;
    std::wstring hotkeyRange = L"-";
    if (hasCandidates) {
        const size_t maxHotkey = std::min(static_cast<size_t>(9), candidates_.size());
        if (maxHotkey <= 1) {
            hotkeyRange = L"1";
        }
        else {
            hotkeyRange = L"1-" + std::to_wstring(maxHotkey);
        }
    }

    drawFooterPart(L"Select: ", RGB(120, 135, 161), footerX);
    drawFooterPart(hasCandidates ? L"Up/Down Space" : L"Backspace/Esc", hasCandidates ? RGB(72, 101, 154) : RGB(163, 176, 199), footerX);
    drawFooterPart(L"   Tab: ", RGB(120, 135, 161), footerX);
    drawFooterPart(hasCandidates ? L"next/prev" : L"-", hasCandidates ? RGB(72, 101, 154) : RGB(163, 176, 199), footerX);
    drawFooterPart(L"   Enter: ", RGB(120, 135, 161), footerX);
    drawFooterPart(hasCandidates ? L"exact/raw" : L"raw code", RGB(72, 101, 154), footerX);
    drawFooterPart(L"   Hotkeys: ", RGB(120, 135, 161), footerX);
    drawFooterPart(hotkeyRange, hasCandidates ? RGB(72, 101, 154) : RGB(163, 176, 199), footerX);
    drawFooterPart(L"   Page: ", RGB(120, 135, 161), footerX);
    drawFooterPart(L"[ ] , . PgUp/PgDn", (canPrevPage || canNextPage) ? RGB(72, 101, 154) : RGB(163, 176, 199), footerX);
    drawFooterPart(L"   Pin: ", RGB(120, 135, 161), footerX);
    drawFooterPart(L"Ctrl+1-9", hasCandidates ? RGB(72, 101, 154) : RGB(163, 176, 199), footerX);
    drawFooterPart(L"   Delete: ", RGB(120, 135, 161), footerX);
    drawFooterPart(L"Ctrl+Del", hasCandidates ? RGB(72, 101, 154) : RGB(163, 176, 199), footerX);
    drawFooterPart(L"   Menu: ", RGB(120, 135, 161), footerX);
    drawFooterPart(L"F2", RGB(72, 101, 154), footerX);

    SelectObject(hdc, oldFont);

    EndPaint(hwnd_, &ps);
}
