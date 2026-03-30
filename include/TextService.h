#pragma once

#include <Windows.h>
#include <msctf.h>

#include "CandidateWindow.h"
#include "CompositionEngine.h"

#include <deque>
#include <filesystem>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>

class TextService final : public ITfTextInputProcessor, public ITfKeyEventSink, public ITfThreadMgrEventSink {
public:
    TextService();

    // IUnknown
    STDMETHODIMP QueryInterface(REFIID riid, void** ppvObj) override;
    STDMETHODIMP_(ULONG) AddRef() override;
    STDMETHODIMP_(ULONG) Release() override;

    // ITfTextInputProcessor
    STDMETHODIMP Activate(ITfThreadMgr* threadMgr, TfClientId clientId) override;
    STDMETHODIMP Deactivate() override;

    // ITfKeyEventSink
    STDMETHODIMP OnSetFocus(BOOL foreground) override;
    STDMETHODIMP OnTestKeyDown(ITfContext* context, WPARAM wParam, LPARAM lParam, BOOL* eaten) override;
    STDMETHODIMP OnTestKeyUp(ITfContext* context, WPARAM wParam, LPARAM lParam, BOOL* eaten) override;
    STDMETHODIMP OnKeyDown(ITfContext* context, WPARAM wParam, LPARAM lParam, BOOL* eaten) override;
    STDMETHODIMP OnKeyUp(ITfContext* context, WPARAM wParam, LPARAM lParam, BOOL* eaten) override;
    STDMETHODIMP OnPreservedKey(ITfContext* context, REFGUID guid, BOOL* eaten) override;

    // ITfThreadMgrEventSink
    STDMETHODIMP OnInitDocumentMgr(ITfDocumentMgr* documentMgr) override;
    STDMETHODIMP OnUninitDocumentMgr(ITfDocumentMgr* documentMgr) override;
    STDMETHODIMP OnSetFocus(ITfDocumentMgr* documentMgrFocus, ITfDocumentMgr* documentMgrPrevFocus) override;
    STDMETHODIMP OnPushContext(ITfContext* context) override;
    STDMETHODIMP OnPopContext(ITfContext* context) override;

private:
    struct CandidateItem {
        std::wstring text;
        std::wstring code;
        std::wstring commitCode;
        std::uint64_t contextScore = 0;
        bool exactMatch = false;
        bool boostedUser = false;
        bool boostedLearned = false;
        bool boostedContext = false;
        bool fromAutoPhrase = false;
        bool fromSystemDict = false;
        bool boostedAutoRepeat = false;
        size_t consumedLength = 0;
    };

    struct CommitHistoryItem {
        std::wstring code;
        std::wstring text;
        ULONGLONG tick = 0;
    };

    struct PendingPhraseStat {
        int hitCount = 0;
        ULONGLONG lastTick = 0;
    };

    enum class ToggleHotkey {
        F8,
        F9,
        CtrlSpace,
    };

    enum class DictionaryProfile {
        ZhengmaLarge,
        ZhengmaLargePinyin,
    };

    static constexpr size_t kDefaultPageSize = 6;

    ~TextService();

    void RefreshCandidates();
    bool ReloadActiveDictionaries();
    bool LoadConfiguredDictionaries();
    bool SwitchDictionaryProfile(DictionaryProfile profile);
    static const wchar_t* GetDictionaryProfileName(DictionaryProfile profile);
    void UpdateCandidateWindow();
    const std::vector<CandidateWindow::DisplayCandidate>& GetCurrentPageCandidates() const;
    void InvalidatePageCandidatesCache();
    size_t GetTotalPages() const;
    void ClearComposition();
    void LoadSettings();
    bool PersistDictionaryProfileSetting() const;
    void NotifyDictionaryProfileSwitch(DictionaryProfile profile, bool success) const;
    bool IsToggleHotkeyPressed(WPARAM wParam) const;
    bool PromoteSelectedCandidateToManualEntry();
    void AppendPhraseReviewEntry(const std::wstring& code, const std::wstring& text, const wchar_t* sourceTag) const;
    void EnsureSingleCharZhengmaCodeHintsLoaded(const std::filesystem::path& dataDir);
    std::wstring GetSingleCharZhengmaCodeHint(const std::wstring& text) const;
    bool CommitCandidateByGlobalIndex(ITfContext* context, size_t globalIndex, std::uint64_t freqBoost);
    bool PinCandidateByGlobalIndex(size_t globalIndex);
    bool BlockCandidateByGlobalIndex(size_t globalIndex);
    bool ShowStatusMenu();
    void LearnPhraseFromRecentCommits(const std::wstring& committedCode, const std::wstring& committedText);
    size_t FindPreferredSelectionIndexForPage(size_t pageIndex) const;
    bool TryFindExactCommitCandidateIndex(size_t& outIndex) const;
    bool TryFindUniqueExactCommitCandidateIndex(size_t& outIndex) const;
    bool LoadContextAssociationFromFile(const std::wstring& filePath);
    bool SaveContextAssociationToFile(const std::wstring& filePath) const;
    bool LoadContextAssociationBlacklistFromFile(const std::wstring& filePath);
    bool SaveContextAssociationBlacklistToFile(const std::wstring& filePath) const;
    void RecordContextAssociation(const std::wstring& prevText, const std::wstring& nextText, std::uint64_t boost);
    std::uint64_t QueryContextAssociationScore(const std::wstring& prevText, const std::wstring& nextText) const;
    static std::wstring MakeContextAssociationKey(const std::wstring& prevText, const std::wstring& nextText);

    bool CommitText(ITfContext* context, const std::wstring& text);
    bool CommitAsciiKey(ITfContext* context, WPARAM wParam, LPARAM lParam);

    static bool IsAlphaKey(WPARAM wParam);
    static wchar_t ToLowerAlpha(WPARAM wParam);

    LONG refCount_;
    ITfThreadMgr* threadMgr_;
    TfClientId clientId_;
    ITfLangBarItemMgr* langBarItemMgr_;
    ITfLangBarItem* configLangBarItem_;
    bool keyEventSinkAdvised_;
    bool threadMgrEventSinkAdvised_;
    DWORD threadMgrEventSinkCookie_;

    CompositionEngine engine_;
    CandidateWindow candidateWindow_;

    bool chineseMode_;
    bool fullShapeMode_;
    bool chinesePunctuation_;
    bool smartSymbolPairs_;
    bool autoCommitUniqueExact_;
    int autoCommitMinCodeLength_;
    bool emptyCandidateBeep_;
    bool tabNavigation_;
    bool enterExactPriority_;
    bool contextAssociationEnabled_;
    int contextAssociationMaxEntries_;
    DictionaryProfile dictionaryProfile_;
    bool nextSingleQuoteOpen_;
    bool nextDoubleQuoteOpen_;
    bool leftShiftTogglePending_;
    ULONGLONG leftShiftToggleDownTick_;
    size_t pageSize_;
    ToggleHotkey toggleHotkey_;

    std::wstring userDataDir_;
    std::wstring userFreqPath_;
    std::wstring userDictPath_;
    std::wstring autoPhraseDictPath_;
    std::wstring blockedEntriesPath_;
    std::wstring contextAssocPath_;
    std::wstring contextAssocBlacklistPath_;
    std::wstring manualPhraseReviewPath_;

    std::wstring compositionCode_;
    std::vector<CandidateItem> allCandidates_;
    mutable std::vector<CandidateWindow::DisplayCandidate> pageCandidatesCache_;
    mutable std::wstring pageCandidatesCacheCompositionCode_;
    mutable size_t pageCandidatesCachePageIndex_;
    mutable size_t pageCandidatesCachePageSize_;
    mutable std::uint64_t pageCandidatesCacheRevision_;
    std::uint64_t candidatesRevision_;
    mutable bool pageCandidatesCacheValid_;
    size_t pageIndex_;
    size_t selectedIndexInPage_;
    bool emptyCandidateAlerted_;
    bool hasRecentAnchor_;
    POINT lastAnchor_;
    ULONGLONG lastAnchorTick_;
    std::deque<CommitHistoryItem> recentCommits_;
    std::unordered_map<std::wstring, PendingPhraseStat> pendingPhraseStats_;
    std::wstring lastAutoPhraseSelectedKey_;
    int autoPhraseSelectedStreak_;
    ULONGLONG autoPhraseSelectedTick_;
    std::unordered_map<std::wstring, std::uint64_t> contextAssociationScores_;
    std::unordered_set<std::wstring> contextAssociationBlacklist_;
    std::unordered_map<wchar_t, std::wstring> singleCharZhengmaCodeHints_;
};

HRESULT CreateTextServiceClassFactory(REFIID riid, void** ppv);
