#pragma once

#include <Windows.h>
#include <msctf.h>

#include "CandidateWindow.h"
#include "CompositionEngine.h"

#include <condition_variable>
#include <deque>
#include <filesystem>
#include <mutex>
#include <string>
#include <thread>
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
        std::uint64_t learnedScore = 0;
        bool exactMatch = false;
        bool boostedUser = false;
        bool boostedLearned = false;
        bool boostedContext = false;
        bool fromAutoPhrase = false;
        bool fromSessionAutoPhrase = false;
        bool fromSystemDict = false;
        bool boostedAutoRepeat = false;
        bool sortSingleChar = false;
        bool sortAutoOnly = false;
        bool sortSystemFiveCodePhrase = false;
        bool sortGB2312Text = false;
        bool sortNonGB2312Single = false;
        bool sortOneCodeSingle = false;
        bool sortOneCodeSingleUsed = false;
        bool sortTwoCodeSingleOrPhrase = false;
        std::uint8_t sortPrimaryTier = 7;
        std::uint8_t sortShortCodeTier = 3;
        size_t consumedLength = 0;
    };

    struct SessionAutoPhraseEntry {
        std::wstring text;
        std::vector<std::wstring> codes;
        ULONGLONG lastTick = 0;
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

    struct PendingAsyncFileWrite {
        std::wstring path;
        std::string content;
        bool deleteIfEmpty = false;
        bool append = false;
        std::uint64_t generation = 0;
        std::uint64_t completedGeneration = 0;
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
    size_t GetCurrentPageCandidateCount() const;
    void InvalidatePageCandidatesCache();
    size_t GetTotalPages() const;
    void ClearComposition();
    void LoadSettings();
    bool PersistDictionaryProfileSetting() const;
    void NotifyDictionaryProfileSwitch(DictionaryProfile profile, bool success) const;
    bool IsToggleHotkeyPressed(WPARAM wParam) const;
    bool PromoteSelectedCandidateToManualEntry();
    void AppendPhraseReviewEntry(const std::wstring& code, const std::wstring& text, const wchar_t* sourceTag);
    void EnsureSingleCharZhengmaCodeHintsLoaded(const std::filesystem::path& dataDir);
    std::wstring GetSingleCharZhengmaCodeHint(const std::wstring& text) const;
    void SyncUserDataFilesStamp();
    bool ReloadUserDataIfChanged(bool force);
    void MarkAutoPhraseDictionaryDirty();
    void MarkFrequencyDataDirty();
    void FlushPendingUserDataIfNeeded(bool force);
    void StartUserDataWriteWorker();
    void StopUserDataWriteWorker(bool waitForPending);
    void UserDataWriteWorkerMain();
    void QueueAutoPhraseSessionWrite();
    void QueueManualPhraseReviewAppend(const std::string& line);
    bool WaitForUserDataWrites(bool includeSessionWrite);
    static bool WriteUtf8FileSnapshot(const std::wstring& filePath, const std::string& content, bool deleteIfEmpty, bool append);
    std::string BuildAutoPhraseSessionStateSnapshot(bool& deleteFile) const;
    std::string BuildContextAssociationFileContent() const;
    bool SaveAutoPhraseSessionState() const;
    bool LoadAutoPhraseSessionState();
    void UpdateSessionAutoPhraseHistory(const std::wstring& committedText, ULONGLONG now);
    void RecordSessionAutoPhraseBreak();
    void CollectSessionAutoPhraseCandidatesForTail(ULONGLONG now);
    void MergeSessionAutoPhraseCandidates();
    bool PromoteSessionAutoPhrase(const std::wstring& text);
    static bool IsHanCharacter(wchar_t ch);
    bool IsTextInGB2312Cached(const std::wstring& text) const;
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
    CompositionEngine phraseBuildEngine_;
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
    int pageBoundaryDirection_;
    int pageBoundaryHitCount_;
    size_t pageSize_;
    ToggleHotkey toggleHotkey_;

    std::wstring userDataDir_;
    std::wstring userFreqPath_;
    std::wstring userDictPath_;
    std::wstring autoPhraseDictPath_;
    std::wstring autoPhraseSessionPath_;
    std::wstring blockedEntriesPath_;
    std::wstring contextAssocPath_;
    std::wstring contextAssocBlacklistPath_;
    std::wstring manualPhraseReviewPath_;
    std::wstring userDataFilesStamp_;

    std::wstring compositionCode_;
    std::vector<CandidateItem> allCandidates_;
    std::uint64_t candidatesRevision_;
    size_t pageIndex_;
    size_t selectedIndexInPage_;
    bool emptyCandidateAlerted_;
    bool hasRecentAnchor_;
    POINT lastAnchor_;
    ULONGLONG lastAnchorTick_;
    std::deque<CommitHistoryItem> recentCommits_;
    std::unordered_map<std::wstring, PendingPhraseStat> pendingPhraseStats_;
    std::wstring autoPhraseHistoryText_;
    std::unordered_map<std::wstring, SessionAutoPhraseEntry> sessionAutoPhraseEntries_;
    std::wstring lastAutoPhraseSelectedKey_;
    int autoPhraseSelectedStreak_;
    ULONGLONG autoPhraseSelectedTick_;
    bool autoPhraseDictionaryDirty_;
    bool userFrequencyDirty_;
    ULONGLONG userDataFirstDirtyTick_;
    ULONGLONG userDataLastFlushTick_;
    std::thread userDataWriteThread_;
    std::mutex userDataWriteMutex_;
    std::condition_variable userDataWriteCv_;
    bool userDataWriteStopRequested_;
    std::uint64_t nextUserDataWriteGeneration_;
    PendingAsyncFileWrite pendingUserDictWrite_;
    PendingAsyncFileWrite pendingUserFreqWrite_;
    PendingAsyncFileWrite pendingContextAssocWrite_;
    PendingAsyncFileWrite pendingBlockedEntriesWrite_;
    PendingAsyncFileWrite pendingAutoPhraseSessionWrite_;
    PendingAsyncFileWrite pendingManualPhraseReviewWrite_;
    std::unordered_map<std::wstring, std::uint64_t> contextAssociationScores_;
    std::unordered_set<std::wstring> contextAssociationBlacklist_;
    std::unordered_map<wchar_t, std::wstring> singleCharZhengmaCodeHints_;
    mutable std::unordered_map<std::wstring, bool> gb2312TextCache_;
    mutable std::uint64_t pageCandidatesCacheRevision_;
    mutable size_t pageCandidatesCachePageIndex_;
    mutable size_t pageCandidatesCachePageSize_;
    mutable std::wstring pageCandidatesCacheCode_;
    mutable std::vector<CandidateWindow::DisplayCandidate> pageCandidatesCache_;
};

HRESULT CreateTextServiceClassFactory(REFIID riid, void** ppv);
