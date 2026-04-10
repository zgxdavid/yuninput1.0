#pragma once

#include <Windows.h>

#include <functional>
#include <string>
#include <vector>

class CandidateWindow {
public:
    struct DisplayCandidate {
        std::wstring text;
        std::wstring code;
        bool boostedUser = false;
        bool boostedLearned = false;
        bool boostedContext = false;
        bool fromAutoPhrase = false;
        bool fromSessionAutoPhrase = false;
        size_t consumedLength = 0;
    };

    CandidateWindow();
    ~CandidateWindow();

    bool EnsureCreated();
    void Destroy();

    void Update(
        const std::wstring& code,
        const std::vector<DisplayCandidate>& candidates,
        size_t pageIndex,
        size_t totalPages,
        size_t totalCandidateCount,
        size_t selectedIndex,
        size_t selectedAbsoluteIndex,
        bool chineseMode,
        bool fullShapeMode,
        const POINT* anchorScreenPos);
    void SetAsyncPollCallback(std::function<void()> callback);
    void ScheduleAsyncPoll(UINT delayMs);
    void CancelAsyncPoll();
    bool IsVisible() const;
    bool Hide();

private:
    static LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam);
    void OnPaint();
    bool EnsureFonts();
    void ReleaseFonts();

    HWND hwnd_;
    std::wstring code_;
    std::vector<DisplayCandidate> candidates_;
    size_t pageIndex_;
    size_t totalPages_;
    size_t totalCandidateCount_;
    size_t selectedIndex_;
    size_t selectedAbsoluteIndex_;
    bool chineseMode_;
    bool fullShapeMode_;
    int displayedQualityLevel_;
    ULONGLONG qualityLevelTick_;
    ULONGLONG qualityPulseUntilTick_;
    bool hasLastWindowRect_;
    int lastX_;
    int lastY_;
    int lastWidth_;
    int lastHeight_;
    size_t lastMeasuredRowCount_;
    size_t lastMeasuredTotalPages_;
    size_t lastMeasuredTotalCandidateCount_;
    HFONT titleFont_;
    HFONT textFont_;
    HFONT selectedTextFont_;
    HFONT smallFont_;
    HFONT codeFont_;
    std::function<void()> asyncPollCallback_;
};
