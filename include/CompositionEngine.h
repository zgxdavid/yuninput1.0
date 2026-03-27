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
        size_t loadOrder = 0;
        bool isUser = false;
        bool isLearned = false;
        bool isAutoPhrase = false;
    };

    struct PhraseCodePart {
        bool fromEnd = false;
        size_t charIndex = 0;
        size_t codeIndex = 0;
        bool noOp = false;
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
    bool LoadFrequencyFromFile(const std::wstring& filePath);
    bool LoadBlockedEntriesFromFile(const std::wstring& filePath);
    bool SaveFrequencyToFile(const std::wstring& filePath) const;
    bool SaveUserDictionaryToFile(const std::wstring& filePath) const;
    bool SaveAutoPhraseDictionaryToFile(const std::wstring& filePath) const;
    bool SaveBlockedEntriesToFile(const std::wstring& filePath) const;
    bool AddUserEntry(const std::wstring& code, const std::wstring& text);
    bool AddAutoPhraseEntry(const std::wstring& code, const std::wstring& text);
    bool PinEntry(const std::wstring& code, const std::wstring& text);
    bool BlockEntry(const std::wstring& code, const std::wstring& text);
    bool TryBuildPhraseCode(const std::wstring& text, std::wstring& outCode) const;

    std::vector<Entry> QueryCandidateEntries(const std::wstring& code, size_t maxCandidates) const;
    std::vector<std::wstring> QueryCandidates(const std::wstring& code, size_t maxCandidates) const;
    void RecordCommit(const std::wstring& code, const std::wstring& text, std::uint64_t boost = 1);

private:
    static std::wstring Utf8ToWide(const std::string& input);
    static std::string WideToUtf8(const std::wstring& input);
    static std::wstring MakeFreqKey(const std::wstring& code, const std::wstring& text);
    static std::wstring NormalizeCode(const std::wstring& code);
    static bool TryParsePhraseRuleToken(const std::string& token, PhraseCodePart& outPart);
    static bool TryParsePhraseRuleSpec(const std::string& key, const std::string& value, PhraseRule& outRule);
    void UpsertPhraseRule(const PhraseRule& rule);
    void ProcessMetadataLine(const std::string& line);
    void LoadDictionaryMetadataFromFile(const std::wstring& filePath);
    bool TryGetBestSingleCharCode(wchar_t ch, size_t minLength, std::wstring& outCode) const;
    bool TryBuildPhraseCodeFromConfiguredRules(const std::vector<std::wstring>& charCodes, std::wstring& outCode) const;

    bool LoadDictionaryInternal(const std::wstring& filePath, bool clearExisting, bool isUserSource, bool isAutoPhraseSource);
    void RebuildIndex();

    std::vector<Entry> entries_;
    std::vector<size_t> sortedIndices_;
    std::unordered_map<std::wstring, std::uint64_t> frequency_;
    std::unordered_set<std::wstring> blockedEntries_;
    std::vector<PhraseRule> phraseRules_;
    std::wstring constructPhrasePrefix_;
};
