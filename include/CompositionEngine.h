#pragma once

#include <cstdint>
#include <string>
#include <unordered_map>
#include <unordered_set>
#include <vector>

class CompositionEngine {
public:
    struct Entry {
        std::wstring code;
        std::wstring text;
        std::uint32_t staticScore = 0;
        std::uint64_t learnedScore = 0;
        size_t loadOrder = 0;
        bool isUser = false;
        bool isLearned = false;
        bool isAutoPhrase = false;
    };

    struct CandidateKey {
        std::wstring code;
        std::wstring text;

        bool operator==(const CandidateKey& other) const {
            return code == other.code && text == other.text;
        }
    };

    struct CandidateKeyHash {
        size_t operator()(const CandidateKey& key) const noexcept {
            const size_t codeHash = std::hash<std::wstring>{}(key.code);
            const size_t textHash = std::hash<std::wstring>{}(key.text);
            return codeHash ^ (textHash + 0x9e3779b97f4a7c15ULL + (codeHash << 6) + (codeHash >> 2));
        }
    };

    struct PrefixRange {
        size_t begin = 0;
        size_t end = 0;
    };

    struct QueryCacheKey {
        std::wstring code;
        size_t maxCandidates = 0;

        bool operator==(const QueryCacheKey& other) const {
            return code == other.code && maxCandidates == other.maxCandidates;
        }
    };

    struct QueryCacheKeyHash {
        size_t operator()(const QueryCacheKey& key) const noexcept {
            const size_t codeHash = std::hash<std::wstring>{}(key.code);
            const size_t countHash = std::hash<size_t>{}(key.maxCandidates);
            return codeHash ^ (countHash + 0x9e3779b97f4a7c15ULL + (codeHash << 6) + (codeHash >> 2));
        }
    };

    struct PhraseCodePart {
        bool fromEnd = false;
        size_t charIndex = 0;
        size_t codeIndex = 0;
        bool noOp = false;

        bool operator==(const PhraseCodePart& other) const {
            return fromEnd == other.fromEnd &&
                   charIndex == other.charIndex &&
                   codeIndex == other.codeIndex &&
                   noOp == other.noOp;
        }
    };

    struct PhraseRule {
        bool exactLength = true;
        size_t length = 0;
        std::vector<PhraseCodePart> parts;
    };

    bool LoadDictionaryFromFile(const std::wstring& filePath);
    bool LoadDictionaryDirectory(const std::wstring& directoryPath);
    bool LoadUserDictionaryFromFile(const std::wstring& filePath);
    bool LoadAutoPhraseDictionaryFromFile(const std::wstring& filePath);
    bool LoadDictionaryMetadataOnlyFromFile(const std::wstring& filePath);
    bool LoadFrequencyFromFile(const std::wstring& filePath);
    bool LoadBlockedEntriesFromFile(const std::wstring& filePath);
    bool SaveFrequencyToFile(const std::wstring& filePath) const;
    bool SaveUserDictionaryToFile(const std::wstring& filePath) const;
    std::string BuildFrequencyFileContent() const;
    std::string BuildUserDictionaryFileContent() const;
    std::string BuildAutoPhraseDictionaryFileContent(size_t maxEntries = 0) const;
    std::string BuildBlockedEntriesFileContent() const;
    bool SaveAutoPhraseDictionaryToFile(const std::wstring& filePath) const;
    bool SaveBlockedEntriesToFile(const std::wstring& filePath) const;
    bool AddUserEntry(const std::wstring& code, const std::wstring& text);
    bool AddAutoPhraseEntry(const std::wstring& code, const std::wstring& text);
    bool PinEntry(const std::wstring& code, const std::wstring& text);
    bool BlockEntry(const std::wstring& code, const std::wstring& text);
    bool TryBuildPhraseCode(const std::wstring& text, std::wstring& outCode) const;
    bool TryBuildPhraseCodes(const std::wstring& text, std::vector<std::wstring>& outCodes) const;
    std::unordered_map<wchar_t, std::wstring> BuildSingleCharCodeHintMap() const;
    bool HasEntry(const std::wstring& code, const std::wstring& text) const;

    std::vector<Entry> QueryCandidateEntries(const std::wstring& code, size_t maxCandidates) const;
    std::vector<Entry> QueryCandidateEntriesFast(const std::wstring& code, size_t maxCandidates, size_t scanBudget) const;
    std::vector<std::wstring> QueryCandidates(const std::wstring& code, size_t maxCandidates) const;
    void RecordCommit(const std::wstring& code, const std::wstring& text, std::uint64_t boost = 1);

private:
    static std::wstring Utf8ToWide(const std::string& input);
    static std::string WideToUtf8(const std::wstring& input);
    static CandidateKey MakeCandidateKey(const std::wstring& code, const std::wstring& text);
    static std::wstring NormalizeCode(const std::wstring& code);
    static bool TryParsePhraseRuleToken(const std::string& token, PhraseCodePart& outPart);
    static bool TryParsePhraseRuleSpec(const std::string& key, const std::string& value, PhraseRule& outRule);
    void UpsertPhraseRule(const PhraseRule& rule);
    void ProcessMetadataLine(const std::string& line);
    void LoadDictionaryMetadataFromFile(const std::wstring& filePath);
    bool TryResolvePhrasePattern(size_t textLength, std::vector<PhraseCodePart>& outPattern) const;
    bool TryCollectPhraseCharCodes(const std::wstring& text, const std::vector<PhraseCodePart>& pattern, std::vector<std::wstring>& outCodes) const;
    bool TryGetBestSingleCharCode(wchar_t ch, size_t minLength, std::wstring& outCode) const;
    bool TryGetSingleCharCodeVariants(wchar_t ch, std::vector<std::wstring>& outCodes) const;
    bool TryBuildPhraseCodeFromConfiguredRules(const std::vector<std::wstring>& charCodes, std::wstring& outCode) const;
    std::pair<std::vector<size_t>::const_iterator, std::vector<size_t>::const_iterator> FindCandidateRange(const std::wstring& normalizedCode) const;
    std::vector<Entry> QueryCandidateEntriesInRange(
        const std::wstring& normalizedCode,
        std::vector<size_t>::const_iterator begin,
        std::vector<size_t>::const_iterator end,
        size_t maxCandidates) const;
    int GetCommonCharRankCached(const std::wstring& text) const;
    void InvalidateQueryCache() const;
    bool EntryIndexLess(size_t left, size_t right) const;
    void RebuildPrefixRanges();
    void InsertEntryIntoIndices(size_t index);

    bool LoadDictionaryInternal(const std::wstring& filePath, bool clearExisting, bool isUserSource, bool isAutoPhraseSource);
    void RebuildIndex();

    std::vector<Entry> entries_;
    std::vector<size_t> sortedIndices_;
    std::unordered_map<CandidateKey, std::uint64_t, CandidateKeyHash> frequency_;
    std::unordered_map<std::wstring, std::uint64_t> textFrequency_;
    std::unordered_set<CandidateKey, CandidateKeyHash> blockedEntries_;
    std::vector<size_t> userEntryIndices_;
    std::vector<size_t> autoPhraseEntryIndices_;
    std::unordered_map<std::wstring, PrefixRange> prefixRanges_;
    std::vector<PhraseRule> phraseRules_;
    std::wstring constructPhrasePrefix_;
    mutable std::unordered_map<std::wstring, int> commonCharRankCache_;
    mutable std::unordered_map<QueryCacheKey, std::vector<Entry>, QueryCacheKeyHash> queryCache_;
};
