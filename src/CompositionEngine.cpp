#include "CompositionEngine.h"

#include <Windows.h>

#include <algorithm>
#include <filesystem>
#include <fstream>
#include <sstream>
#include <unordered_map>
#include <utility>

namespace {

struct CandidateScore {
    std::wstring text;
    std::wstring displayCode;
    bool displayCodeExact = false;
    size_t displayCodeLength = 0;
    size_t displayCodeLoadOrder = 0;
    bool exactCode = false;
    std::uint64_t frequency = 0;
    std::uint32_t staticScore = 0;
    size_t shortestCodeLength = 0;
    size_t earliestLoadOrder = 0;
    bool hasUser = false;
    bool hasManualUser = false;
    bool hasAutoPhrase = false;
    bool hasSystemSource = false;
    bool hasLearned = false;
    int commonCharRank = 100000;
    size_t completionDelta = 0;
    int lengthPreferenceScore = 100;
    bool preferredTwoCharPhrase = false;
    int twoCodePriorityTier = 3;
};

bool IsPreferredTwoCharPhraseCandidate(const CandidateScore& score, size_t queryCodeLength) {
    if (queryCodeLength != 2) {
        return false;
    }

    if (score.text.size() != 2) {
        return false;
    }

    if (!score.exactCode) {
        return false;
    }

    if (!score.hasSystemSource) {
        return false;
    }

    if (score.hasAutoPhrase) {
        return false;
    }

    return true;
}

int GetTwoCodePriorityTier(const CandidateScore& score, size_t queryCodeLength) {
    if (queryCodeLength != 2) {
        return 3;
    }

    if (!score.exactCode) {
        return 3;
    }

    if (score.text.size() == 1 && score.displayCodeLength <= 1) {
        return 0;
    }

    if (score.text.size() == 2 &&
        score.hasSystemSource &&
        !score.hasAutoPhrase &&
        score.shortestCodeLength == 2 &&
        score.displayCodeLength == 2) {
        return 1;
    }

    if (score.text.size() == 1 && score.displayCodeLength == 2) {
        return 2;
    }

    if (score.text.size() == 2 && score.hasSystemSource && !score.hasAutoPhrase) {
        return 3;
    }

    return 4;
}

int GetLengthPreferenceScore(size_t queryCodeLength, size_t textLength) {
    if (textLength == 0) {
        return 100;
    }

    if (queryCodeLength <= 2) {
        if (textLength == 1) {
            return 0;
        }
        if (textLength == 2) {
            return 1;
        }
        return 4;
    }

    if (queryCodeLength == 3) {
        if (textLength == 2) {
            return 0;
        }
        if (textLength == 1) {
            return 1;
        }
        if (textLength == 3) {
            return 2;
        }
        return 4;
    }

    // With longer codes, phrase candidates are usually expected before isolated characters.
    if (textLength >= 2 && textLength <= 4) {
        return 0;
    }
    if (textLength == 1) {
        return 2;
    }
    return 3;
}

int GetCommonCharRank(const std::wstring& text) {
    if (text.empty()) {
        return 100000;
    }

    if (text.size() != 1) {
        return 100000 + static_cast<int>(text.size());
    }

    const wchar_t ch = text[0];

    auto getGb2312Rank = [](wchar_t value, int& outRank) {
        char bytes[4] = {};
        BOOL usedDefaultChar = FALSE;
        const int len = WideCharToMultiByte(
            20936,
            WC_NO_BEST_FIT_CHARS,
            &value,
            1,
            bytes,
            static_cast<int>(sizeof(bytes)),
            nullptr,
            &usedDefaultChar);
        if (len != 2 || usedDefaultChar) {
            return false;
        }

        const unsigned char b1 = static_cast<unsigned char>(bytes[0]);
        const unsigned char b2 = static_cast<unsigned char>(bytes[1]);
        if (b1 < 0xA1 || b1 > 0xF7 || b2 < 0xA1 || b2 > 0xFE) {
            return false;
        }

        // GB2312 level-1 (commonly used) first, then level-2.
        if (b1 >= 0xB0 && b1 <= 0xD7) {
            outRank = static_cast<int>(b1 - 0xB0);
            return true;
        }
        if (b1 >= 0xD8 && b1 <= 0xF7) {
            outRank = 80 + static_cast<int>(b1 - 0xD8);
            return true;
        }

        // Symbols/full-width forms in GB2312 are valid but lower priority than Han chars.
        outRank = 180 + static_cast<int>(b1 - 0xA1);
        return true;
    };

    int gb2312Rank = 0;
    if (getGb2312Rank(ch, gb2312Rank)) {
        return gb2312Rank;
    }

    if (ch >= 0x4E00 && ch <= 0x9FFF) {
        // Prefer base unified CJK ideographs first.
        return 320;
    }
    if (ch >= 0x3400 && ch <= 0x4DBF) {
        // Extension A is generally less common than base block.
        return 420;
    }
    if (ch >= 0xF900 && ch <= 0xFAFF) {
        return 520;
    }
    if (ch >= 0x2E80 && ch <= 0x2EFF) {
        return 620;
    }
    if ((ch >= 0x3000 && ch <= 0x303F) || (ch >= 0xFF00 && ch <= 0xFFEF)) {
        return 720;
    }

    return 900;
}

int GetSingleCharShortCodeRank(const CandidateScore& score) {
    if (score.text.size() != 1) {
        return 100;
    }

    if (score.shortestCodeLength <= 1) {
        return 0;
    }
    if (score.shortestCodeLength == 2) {
        return 1;
    }
    if (score.shortestCodeLength == 3) {
        return 2;
    }
    return 3;
}

bool IsCodeToken(const std::string& token) {
    if (token.empty()) {
        return false;
    }

    for (const unsigned char ch : token) {
        if (!((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z'))) {
            return false;
        }
    }

    return true;
}

bool IsLikelyBrokenCandidate(const std::wstring& text) {
    if (text.empty()) {
        return true;
    }

    for (wchar_t ch : text) {
        if (ch == static_cast<wchar_t>(0x20AC) ||
            ch == static_cast<wchar_t>(0x00A4) ||
            ch == static_cast<wchar_t>(0xFFFD) ||
            ch == static_cast<wchar_t>(0xFFFE) ||
            ch == static_cast<wchar_t>(0xFFFF)) {
            return true;
        }
    }

    return false;
}

bool AppendPhraseCodePart(
    const std::vector<std::wstring>& charCodes,
    const CompositionEngine::PhraseCodePart& part,
    std::wstring& outCode) {
    if (charCodes.empty()) {
        return false;
    }

    if (part.noOp) {
        return true;
    }

    const size_t textIndex = part.fromEnd ? (charCodes.size() - 1) : part.charIndex;
    if (textIndex >= charCodes.size()) {
        return false;
    }

    const std::wstring& code = charCodes[textIndex];
    if (part.codeIndex >= code.size()) {
        return false;
    }

    outCode.push_back(code[part.codeIndex]);
    return true;
}

bool BuildPhraseCodeFromPattern(
    const std::vector<std::wstring>& charCodes,
    const std::vector<CompositionEngine::PhraseCodePart>& pattern,
    std::wstring& outCode) {
    outCode.clear();
    outCode.reserve(pattern.size());
    for (const CompositionEngine::PhraseCodePart& part : pattern) {
        if (!AppendPhraseCodePart(charCodes, part, outCode)) {
            outCode.clear();
            return false;
        }
    }

    return !outCode.empty();
}

}  // namespace

std::wstring CompositionEngine::Utf8ToWide(const std::string& input) {
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

std::string CompositionEngine::WideToUtf8(const std::wstring& input) {
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

std::wstring CompositionEngine::MakeFreqKey(const std::wstring& code, const std::wstring& text) {
    return code + L"\t" + text;
}

std::wstring CompositionEngine::NormalizeCode(const std::wstring& code) {
    std::wstring normalized;
    normalized.reserve(code.size());

    for (wchar_t ch : code) {
        if (ch >= L'A' && ch <= L'Z') {
            normalized.push_back(static_cast<wchar_t>(ch - L'A' + L'a'));
        } else {
            normalized.push_back(ch);
        }
    }

    return normalized;
}

bool CompositionEngine::TryParsePhraseRuleToken(const std::string& token, PhraseCodePart& outPart) {
    outPart = PhraseCodePart{};
    if (token.size() != 3) {
        return false;
    }

    const char direction = static_cast<char>(std::tolower(static_cast<unsigned char>(token[0])));
    if (direction != 'p' && direction != 'n') {
        return false;
    }

    if (!std::isdigit(static_cast<unsigned char>(token[1])) || !std::isdigit(static_cast<unsigned char>(token[2]))) {
        return false;
    }

    const int charOrdinal = token[1] - '0';
    const int codeOrdinal = token[2] - '0';
    if (charOrdinal == 0 || codeOrdinal == 0) {
        outPart.noOp = true;
        return true;
    }

    outPart.fromEnd = direction == 'n';
    outPart.charIndex = static_cast<size_t>(charOrdinal - 1);
    outPart.codeIndex = static_cast<size_t>(codeOrdinal - 1);
    return true;
}

bool CompositionEngine::TryParsePhraseRuleSpec(const std::string& key, const std::string& value, PhraseRule& outRule) {
    outRule = PhraseRule{};
    if (key.size() < 2) {
        return false;
    }

    const char scope = static_cast<char>(std::tolower(static_cast<unsigned char>(key[0])));
    if (scope != 'e' && scope != 'a') {
        return false;
    }

    size_t length = 0;
    try {
        length = static_cast<size_t>(std::stoul(key.substr(1)));
    }
    catch (...) {
        return false;
    }

    if (length == 0) {
        return false;
    }

    std::vector<PhraseCodePart> parts;
    std::istringstream iss(value);
    std::string token;
    while (std::getline(iss, token, '+')) {
        PhraseCodePart part;
        if (!TryParsePhraseRuleToken(token, part)) {
            return false;
        }
        parts.push_back(part);
    }

    if (parts.empty()) {
        return false;
    }

    outRule.exactLength = scope == 'e';
    outRule.length = length;
    outRule.parts = std::move(parts);
    return true;
}

void CompositionEngine::UpsertPhraseRule(const PhraseRule& rule) {
    for (PhraseRule& existing : phraseRules_) {
        if (existing.exactLength == rule.exactLength && existing.length == rule.length) {
            existing = rule;
            return;
        }
    }

    phraseRules_.push_back(rule);
    std::sort(
        phraseRules_.begin(),
        phraseRules_.end(),
        [](const PhraseRule& left, const PhraseRule& right) {
            if (left.exactLength != right.exactLength) {
                return left.exactLength > right.exactLength;
            }
            return left.length < right.length;
        });
}

void CompositionEngine::ProcessMetadataLine(const std::string& line) {
    const std::string constructPrefix = "# yuninput:construct_phrase=";
    const std::string rulePrefix = "# yuninput:rule:";
    if (line.rfind(constructPrefix, 0) == 0) {
        constructPhrasePrefix_ = Utf8ToWide(line.substr(constructPrefix.size()));
        return;
    }

    if (line.rfind(rulePrefix, 0) == 0) {
        const std::string spec = line.substr(rulePrefix.size());
        const size_t equalPos = spec.find('=');
        if (equalPos != std::string::npos) {
            PhraseRule rule;
            if (TryParsePhraseRuleSpec(spec.substr(0, equalPos), spec.substr(equalPos + 1), rule)) {
                UpsertPhraseRule(rule);
            }
        }
    }
}

void CompositionEngine::LoadDictionaryMetadataFromFile(const std::wstring& filePath) {
    std::ifstream input(filePath, std::ios::in | std::ios::binary);
    if (!input) {
        return;
    }

    std::string line;
    while (std::getline(input, line)) {
        if (line.empty()) {
            continue;
        }

        ProcessMetadataLine(line);
    }
}

bool CompositionEngine::TryGetBestSingleCharCode(wchar_t ch, size_t minLength, std::wstring& outCode) const {
    outCode.clear();

    const auto isBetterCode = [](const Entry& candidate, const Entry* bestEntry) {
        if (bestEntry == nullptr) {
            return true;
        }
        if (candidate.code.size() != bestEntry->code.size()) {
            return candidate.code.size() < bestEntry->code.size();
        }
        if (candidate.staticScore != bestEntry->staticScore) {
            return candidate.staticScore > bestEntry->staticScore;
        }
        return candidate.loadOrder < bestEntry->loadOrder;
    };

    const Entry* bestEntry = nullptr;
    for (const Entry& entry : entries_) {
        if (entry.text.size() != 1 || entry.text[0] != ch) {
            continue;
        }
        if (entry.code.size() < minLength) {
            continue;
        }
        if (blockedEntries_.find(MakeFreqKey(entry.code, entry.text)) != blockedEntries_.end()) {
            continue;
        }
        if (isBetterCode(entry, bestEntry)) {
            bestEntry = &entry;
        }
    }

    if (bestEntry == nullptr) {
        return false;
    }

    outCode = bestEntry->code;
    return true;
}

bool CompositionEngine::TryBuildPhraseCodeFromConfiguredRules(const std::vector<std::wstring>& charCodes, std::wstring& outCode) const {
    outCode.clear();
    if (charCodes.size() < 2) {
        return false;
    }

    const size_t phraseLength = charCodes.size();
    const PhraseRule* bestAtLeastRule = nullptr;
    for (const PhraseRule& rule : phraseRules_) {
        if (rule.exactLength) {
            if (rule.length == phraseLength) {
                return BuildPhraseCodeFromPattern(charCodes, rule.parts, outCode);
            }
            continue;
        }

        if (rule.length <= phraseLength) {
            if (bestAtLeastRule == nullptr || rule.length > bestAtLeastRule->length) {
                bestAtLeastRule = &rule;
            }
        }
    }

    if (bestAtLeastRule != nullptr) {
        return BuildPhraseCodeFromPattern(charCodes, bestAtLeastRule->parts, outCode);
    }

    return false;
}

bool CompositionEngine::TryBuildPhraseCode(const std::wstring& text, std::wstring& outCode) const {
    outCode.clear();
    if (text.size() < 2) {
        return false;
    }

    std::vector<std::wstring> charCodes;
    charCodes.reserve(text.size());

    size_t requiredCodeLength = 1;
    if (text.size() == 2) {
        requiredCodeLength = 2;
    }
    else if (text.size() == 3) {
        // Zhengma 3-char phrase rule needs each source char to come from 3..4 full code.
        requiredCodeLength = 3;
    }

    for (wchar_t ch : text) {
        std::wstring charCode;
        if (!TryGetBestSingleCharCode(ch, requiredCodeLength, charCode)) {
            return false;
        }
        charCodes.push_back(std::move(charCode));
    }

    static const std::vector<PhraseCodePart> kRule4Plus = {
        {false, 0, 0, false},
        {false, 1, 0, false},
        {false, 2, 0, false},
        {false, 3, 0, false},
    };

    // Native Zhengma 4+ phrase rule: use first-four head codes.
    if (text.size() > 4) {
        return BuildPhraseCodeFromPattern(charCodes, kRule4Plus, outCode);
    }

    if (TryBuildPhraseCodeFromConfiguredRules(charCodes, outCode)) {
        return true;
    }

    static const std::vector<PhraseCodePart> kRule2 = {
        {false, 0, 0, false},
        {false, 0, 1, false},
        {false, 1, 0, false},
        {false, 1, 1, false},
    };
    static const std::vector<PhraseCodePart> kRule3 = {
        {false, 0, 0, false},
        {false, 1, 0, false},
        {false, 1, 2, false},
        {false, 2, 0, false},
    };
    if (text.size() == 2) {
        return BuildPhraseCodeFromPattern(charCodes, kRule2, outCode);
    }
    if (text.size() == 3) {
        return BuildPhraseCodeFromPattern(charCodes, kRule3, outCode);
    }

    return BuildPhraseCodeFromPattern(charCodes, kRule4Plus, outCode);
}

void CompositionEngine::RebuildIndex() {
    sortedIndices_.clear();
    sortedIndices_.reserve(entries_.size());
    for (size_t i = 0; i < entries_.size(); ++i) {
        sortedIndices_.push_back(i);
    }

    std::stable_sort(
        sortedIndices_.begin(),
        sortedIndices_.end(),
        [this](size_t left, size_t right) {
            const Entry& l = entries_[left];
            const Entry& r = entries_[right];
            if (l.code != r.code) {
                return l.code < r.code;
            }
            if (l.staticScore != r.staticScore) {
                return l.staticScore > r.staticScore;
            }
            if (l.loadOrder != r.loadOrder) {
                return l.loadOrder < r.loadOrder;
            }
            return l.text < r.text;
        });
}

bool CompositionEngine::LoadDictionaryInternal(const std::wstring& filePath, bool clearExisting, bool isUserSource, bool isAutoPhraseSource) {
    std::ifstream input(filePath, std::ios::in | std::ios::binary);
    if (!input) {
        return false;
    }

    if (clearExisting) {
        entries_.clear();
        phraseRules_.clear();
        constructPhrasePrefix_.clear();
    }

    LoadDictionaryMetadataFromFile(filePath + L".rules");

    std::string line;
    while (std::getline(input, line)) {
        if (line.empty()) {
            continue;
        }

        if (line[0] == '#') {
            // Metadata is loaded from sidecar .rules files to avoid coupling rule
            // changes to large dictionary-body file rewrites.
            continue;
        }

        std::istringstream iss(line);
        std::string firstToken;
        std::string secondToken;
        std::string thirdToken;
        if (!(iss >> firstToken >> secondToken)) {
            continue;
        }

        const bool firstIsCode = IsCodeToken(firstToken);
        const bool secondIsCode = IsCodeToken(secondToken);

        std::string codeUtf8;
        std::string textUtf8;
        if (firstIsCode && !secondIsCode) {
            codeUtf8 = firstToken;
            textUtf8 = secondToken;
        } else if (!firstIsCode && secondIsCode) {
            textUtf8 = firstToken;
            codeUtf8 = secondToken;
        } else {
            continue;
        }

        std::uint32_t staticScore = 0;
        if (iss >> thirdToken) {
            try {
                staticScore = static_cast<std::uint32_t>(std::stoul(thirdToken));
            }
            catch (...) {
                staticScore = 0;
            }
        }

        Entry entry;
        entry.code = NormalizeCode(Utf8ToWide(codeUtf8));
        entry.text = Utf8ToWide(textUtf8);
        entry.staticScore = staticScore;
        entry.loadOrder = entries_.size();
        entry.isUser = isUserSource;
        entry.isAutoPhrase = isAutoPhraseSource;
        if (!entry.code.empty() && !entry.text.empty()) {
            entries_.push_back(std::move(entry));
        }
    }

    RebuildIndex();

    return !entries_.empty();
}

bool CompositionEngine::LoadDictionaryFromFile(const std::wstring& filePath) {
    return LoadDictionaryInternal(filePath, true, false, false);
}

bool CompositionEngine::LoadDictionaryDirectory(const std::wstring& directoryPath) {
    namespace fs = std::filesystem;

    std::error_code ec;
    if (!fs::exists(directoryPath, ec) || !fs::is_directory(directoryPath, ec)) {
        return false;
    }

    std::vector<fs::path> dictFiles;
    for (const auto& entry : fs::directory_iterator(directoryPath, ec)) {
        if (ec) {
            break;
        }

        if (!entry.is_regular_file()) {
            continue;
        }

        if (_wcsicmp(entry.path().extension().wstring().c_str(), L".dict") != 0) {
            continue;
        }

        if (_wcsicmp(entry.path().filename().wstring().c_str(), L"yuninput_user.dict") == 0) {
            continue;
        }

        dictFiles.push_back(entry.path());
    }

    std::sort(dictFiles.begin(), dictFiles.end());

    bool anyLoaded = false;
    bool firstFile = true;
    for (const auto& path : dictFiles) {
        if (LoadDictionaryInternal(path.wstring(), firstFile, false, false)) {
            anyLoaded = true;
            firstFile = false;
        }
    }

    return anyLoaded;
}

bool CompositionEngine::LoadUserDictionaryFromFile(const std::wstring& filePath) {
    return LoadDictionaryInternal(filePath, false, true, false);
}

bool CompositionEngine::LoadAutoPhraseDictionaryFromFile(const std::wstring& filePath) {
    return LoadDictionaryInternal(filePath, false, true, true);
}

bool CompositionEngine::LoadFrequencyFromFile(const std::wstring& filePath) {
    std::ifstream input(filePath, std::ios::in | std::ios::binary);
    if (!input) {
        return false;
    }

    frequency_.clear();

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
            continue;
        }

        const std::wstring code = Utf8ToWide(codeUtf8);
        const std::wstring text = Utf8ToWide(textUtf8);
        if (code.empty() || text.empty()) {
            continue;
        }

        frequency_[MakeFreqKey(code, text)] = score;
    }

    return true;
}

bool CompositionEngine::LoadBlockedEntriesFromFile(const std::wstring& filePath) {
    std::ifstream input(filePath, std::ios::in | std::ios::binary);
    if (!input) {
        return false;
    }

    blockedEntries_.clear();

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

        const std::wstring code = NormalizeCode(Utf8ToWide(codeUtf8));
        const std::wstring text = Utf8ToWide(textUtf8);
        if (code.empty() || text.empty()) {
            continue;
        }

        blockedEntries_.insert(MakeFreqKey(code, text));
    }

    return true;
}

bool CompositionEngine::SaveFrequencyToFile(const std::wstring& filePath) const {
    std::ofstream output(filePath, std::ios::out | std::ios::binary | std::ios::trunc);
    if (!output) {
        return false;
    }

    for (const auto& pair : frequency_) {
        const size_t split = pair.first.find(L'\t');
        if (split == std::wstring::npos) {
            continue;
        }

        const std::wstring code = pair.first.substr(0, split);
        const std::wstring text = pair.first.substr(split + 1);
        output << WideToUtf8(code) << ' ' << WideToUtf8(text) << ' ' << pair.second << '\n';
    }

    return true;
}

bool CompositionEngine::SaveBlockedEntriesToFile(const std::wstring& filePath) const {
    std::ofstream output(filePath, std::ios::out | std::ios::binary | std::ios::trunc);
    if (!output) {
        return false;
    }

    for (const auto& blockedKey : blockedEntries_) {
        const size_t split = blockedKey.find(L'\t');
        if (split == std::wstring::npos) {
            continue;
        }

        const std::wstring code = blockedKey.substr(0, split);
        const std::wstring text = blockedKey.substr(split + 1);
        output << WideToUtf8(code) << ' ' << WideToUtf8(text) << '\n';
    }

    return true;
}

bool CompositionEngine::SaveUserDictionaryToFile(const std::wstring& filePath) const {
    std::ofstream output(filePath, std::ios::out | std::ios::binary | std::ios::trunc);
    if (!output) {
        return false;
    }

    for (const Entry& entry : entries_) {
        if (!entry.isUser || entry.isAutoPhrase || entry.code.empty() || entry.text.empty()) {
            continue;
        }

        output << WideToUtf8(entry.code) << ' ' << WideToUtf8(entry.text) << ' ' << entry.staticScore << '\n';
    }

    return true;
}

bool CompositionEngine::SaveAutoPhraseDictionaryToFile(const std::wstring& filePath) const {
    std::ofstream output(filePath, std::ios::out | std::ios::binary | std::ios::trunc);
    if (!output) {
        return false;
    }

    for (const Entry& entry : entries_) {
        if (!entry.isAutoPhrase || entry.code.empty() || entry.text.empty()) {
            continue;
        }

        output << WideToUtf8(entry.code) << ' ' << WideToUtf8(entry.text) << ' ' << entry.staticScore << '\n';
    }

    return true;
}

bool CompositionEngine::AddUserEntry(const std::wstring& code, const std::wstring& text) {
    const std::wstring normalizedCode = NormalizeCode(code);
    if (normalizedCode.empty() || text.empty()) {
        return false;
    }

    blockedEntries_.erase(MakeFreqKey(normalizedCode, text));

    for (Entry& entry : entries_) {
        if (entry.code == normalizedCode && entry.text == text) {
            entry.isUser = true;
            entry.isAutoPhrase = false;
            return false;
        }
    }

    Entry entry;
    entry.code = normalizedCode;
    entry.text = text;
    entry.staticScore = 1;
    entry.loadOrder = entries_.size();
    entry.isUser = true;
    entry.isAutoPhrase = false;
    entries_.push_back(std::move(entry));
    RebuildIndex();
    return true;
}

bool CompositionEngine::AddAutoPhraseEntry(const std::wstring& code, const std::wstring& text) {
    const std::wstring normalizedCode = NormalizeCode(code);
    if (normalizedCode.empty() || text.empty()) {
        return false;
    }

    blockedEntries_.erase(MakeFreqKey(normalizedCode, text));

    for (Entry& entry : entries_) {
        if (entry.code == normalizedCode && entry.text == text) {
            if (entry.isUser && !entry.isAutoPhrase) {
                return false;
            }
            const bool changed = !entry.isAutoPhrase;
            entry.isUser = true;
            entry.isAutoPhrase = true;
            return changed;
        }
    }

    Entry entry;
    entry.code = normalizedCode;
    entry.text = text;
    entry.staticScore = 1;
    entry.loadOrder = entries_.size();
    entry.isUser = true;
    entry.isAutoPhrase = true;
    entries_.push_back(std::move(entry));
    RebuildIndex();
    return true;
}

bool CompositionEngine::PinEntry(const std::wstring& code, const std::wstring& text) {
    const std::wstring normalizedCode = NormalizeCode(code);
    if (normalizedCode.empty() || text.empty()) {
        return false;
    }

    blockedEntries_.erase(MakeFreqKey(normalizedCode, text));

    for (Entry& entry : entries_) {
        if (entry.code == normalizedCode && entry.text == text) {
            const bool changed = !entry.isUser;
            entry.isUser = true;
            entry.isAutoPhrase = false;
            return changed;
        }
    }

    Entry entry;
    entry.code = normalizedCode;
    entry.text = text;
    entry.staticScore = 1;
    entry.loadOrder = entries_.size();
    entry.isUser = true;
    entry.isAutoPhrase = false;
    entries_.push_back(std::move(entry));
    RebuildIndex();
    return true;
}

bool CompositionEngine::BlockEntry(const std::wstring& code, const std::wstring& text) {
    const std::wstring normalizedCode = NormalizeCode(code);
    if (normalizedCode.empty() || text.empty()) {
        return false;
    }

    const std::wstring key = MakeFreqKey(normalizedCode, text);
    const bool inserted = blockedEntries_.insert(key).second;
    frequency_.erase(key);

    const size_t oldSize = entries_.size();
    entries_.erase(
        std::remove_if(
            entries_.begin(),
            entries_.end(),
            [&](const Entry& entry) {
                return entry.code == normalizedCode && entry.text == text && entry.isUser;
            }),
        entries_.end());

    if (entries_.size() != oldSize) {
        for (size_t i = 0; i < entries_.size(); ++i) {
            entries_[i].loadOrder = i;
        }
        RebuildIndex();
    }

    return inserted || entries_.size() != oldSize;
}

std::vector<CompositionEngine::Entry> CompositionEngine::QueryCandidateEntries(const std::wstring& code, size_t maxCandidates) const {
    std::vector<Entry> result;
    if (code.empty() || maxCandidates == 0 || sortedIndices_.empty()) {
        return result;
    }

    const std::wstring normalizedCode = NormalizeCode(code);

    const std::wstring rangeEnd = normalizedCode + std::wstring(1, static_cast<wchar_t>(0xFFFF));
    const auto begin = std::lower_bound(
        sortedIndices_.begin(),
        sortedIndices_.end(),
        normalizedCode,
        [this](size_t idx, const std::wstring& key) {
            return entries_[idx].code < key;
        });

    const auto end = std::lower_bound(
        begin,
        sortedIndices_.end(),
        rangeEnd,
        [this](size_t idx, const std::wstring& key) {
            return entries_[idx].code < key;
        });

    std::unordered_map<std::wstring, CandidateScore> bestScoreByText;
    bestScoreByText.reserve(static_cast<size_t>(std::distance(begin, end)));

    for (auto it = begin; it != end; ++it) {
        const Entry& item = entries_[*it];
        const std::wstring freqKey = MakeFreqKey(item.code, item.text);
        if (blockedEntries_.find(freqKey) != blockedEntries_.end()) {
            continue;
        }
        if (IsLikelyBrokenCandidate(item.text)) {
            continue;
        }
        const auto freqIt = frequency_.find(freqKey);
        const std::uint64_t score = (freqIt == frequency_.end()) ? 0 : freqIt->second;
        const bool exactCode = item.code == normalizedCode;
        const size_t completionDelta = item.code.size() >= normalizedCode.size()
            ? (item.code.size() - normalizedCode.size())
            : 0;
        const int lengthPreferenceScore = GetLengthPreferenceScore(normalizedCode.size(), item.text.size());

        auto scoreIt = bestScoreByText.find(item.text);
        if (scoreIt == bestScoreByText.end()) {
            CandidateScore candidate;
            candidate.text = item.text;
            candidate.displayCode = item.code;
            candidate.displayCodeExact = exactCode;
            candidate.displayCodeLength = item.code.size();
            candidate.displayCodeLoadOrder = item.loadOrder;
            candidate.exactCode = exactCode;
            candidate.frequency = score;
            candidate.staticScore = item.staticScore;
            candidate.shortestCodeLength = item.code.size();
            candidate.earliestLoadOrder = item.loadOrder;
            candidate.hasUser = item.isUser;
            candidate.hasManualUser = item.isUser && !item.isAutoPhrase;
            candidate.hasAutoPhrase = item.isAutoPhrase;
            candidate.hasSystemSource = !item.isUser;
            candidate.hasLearned = score > 0;
            candidate.commonCharRank = GetCommonCharRank(item.text);
            candidate.completionDelta = completionDelta;
            candidate.lengthPreferenceScore = lengthPreferenceScore;
            candidate.preferredTwoCharPhrase = IsPreferredTwoCharPhraseCandidate(candidate, normalizedCode.size());
            candidate.twoCodePriorityTier = GetTwoCodePriorityTier(candidate, normalizedCode.size());
            bestScoreByText.emplace(item.text, candidate);
            continue;
        }

        CandidateScore& existing = scoreIt->second;
        existing.exactCode = existing.exactCode || exactCode;
        existing.hasUser = existing.hasUser || item.isUser;
        existing.hasManualUser = existing.hasManualUser || (item.isUser && !item.isAutoPhrase);
        existing.hasAutoPhrase = existing.hasAutoPhrase || item.isAutoPhrase;
        existing.hasSystemSource = existing.hasSystemSource || !item.isUser;
        existing.hasLearned = existing.hasLearned || (score > 0);
        if (score > existing.frequency) {
            existing.frequency = score;
        }
        if (item.staticScore > existing.staticScore) {
            existing.staticScore = item.staticScore;
        }
        if (existing.shortestCodeLength == 0 || item.code.size() < existing.shortestCodeLength) {
            existing.shortestCodeLength = item.code.size();
        }
        if (item.loadOrder < existing.earliestLoadOrder) {
            existing.earliestLoadOrder = item.loadOrder;
        }
        const int rank = GetCommonCharRank(item.text);
        if (rank < existing.commonCharRank) {
            existing.commonCharRank = rank;
        }
        if (completionDelta < existing.completionDelta) {
            existing.completionDelta = completionDelta;
        }
        if (lengthPreferenceScore < existing.lengthPreferenceScore) {
            existing.lengthPreferenceScore = lengthPreferenceScore;
        }
        if (IsPreferredTwoCharPhraseCandidate(existing, normalizedCode.size()) ||
            (item.text.size() == 2 && exactCode && !item.isUser && !item.isAutoPhrase && normalizedCode.size() == 2)) {
            existing.preferredTwoCharPhrase = true;
        }
        existing.twoCodePriorityTier = GetTwoCodePriorityTier(existing, normalizedCode.size());

        const bool betterDisplayCode =
            existing.displayCode.empty() ||
            (exactCode && !existing.displayCodeExact) ||
            (exactCode == existing.displayCodeExact &&
                (item.code.size() < existing.displayCodeLength ||
                    (item.code.size() == existing.displayCodeLength && item.loadOrder < existing.displayCodeLoadOrder)));
        if (betterDisplayCode) {
            existing.displayCode = item.code;
            existing.displayCodeExact = exactCode;
            existing.displayCodeLength = item.code.size();
            existing.displayCodeLoadOrder = item.loadOrder;
        }
    }

    std::vector<CandidateScore> ranked;
    ranked.reserve(bestScoreByText.size());
    for (const auto& pair : bestScoreByText) {
        ranked.push_back(pair.second);
    }

    const size_t queryCodeLength = normalizedCode.size();

    std::stable_sort(
        ranked.begin(),
        ranked.end(),
        [queryCodeLength](const auto& left, const auto& right) {
            const int leftShortCodeRank = GetSingleCharShortCodeRank(left);
            const int rightShortCodeRank = GetSingleCharShortCodeRank(right);
            const bool leftShortFullCode = left.shortestCodeLength < 4;
            const bool rightShortFullCode = right.shortestCodeLength < 4;
            if (left.twoCodePriorityTier != right.twoCodePriorityTier) {
                return left.twoCodePriorityTier < right.twoCodePriorityTier;
            }
            if (left.exactCode != right.exactCode) {
                return left.exactCode > right.exactCode;
            }
            if (left.preferredTwoCharPhrase != right.preferredTwoCharPhrase) {
                return left.preferredTwoCharPhrase > right.preferredTwoCharPhrase;
            }
            if (leftShortFullCode != rightShortFullCode) {
                return leftShortFullCode > rightShortFullCode;
            }
            if (leftShortCodeRank != rightShortCodeRank) {
                return leftShortCodeRank < rightShortCodeRank;
            }
            const bool leftAutoOnly = left.hasAutoPhrase && !left.hasSystemSource && !left.hasManualUser;
            const bool rightAutoOnly = right.hasAutoPhrase && !right.hasSystemSource && !right.hasManualUser;
            if (leftAutoOnly != rightAutoOnly) {
                return leftAutoOnly < rightAutoOnly;
            }
            if (left.hasManualUser != right.hasManualUser) {
                return left.hasManualUser > right.hasManualUser;
            }
            if (left.frequency != right.frequency) {
                return left.frequency > right.frequency;
            }
            if (left.shortestCodeLength != right.shortestCodeLength) {
                // Shorter code level should come first when user/learned signals tie.
                return left.shortestCodeLength < right.shortestCodeLength;
            }
            if (left.commonCharRank != right.commonCharRank) {
                return left.commonCharRank < right.commonCharRank;
            }
            if (left.staticScore != right.staticScore) {
                return left.staticScore > right.staticScore;
            }
            if (left.completionDelta != right.completionDelta) {
                return left.completionDelta < right.completionDelta;
            }
            if (left.lengthPreferenceScore != right.lengthPreferenceScore) {
                return left.lengthPreferenceScore < right.lengthPreferenceScore;
            }
            if (left.earliestLoadOrder != right.earliestLoadOrder) {
                return left.earliestLoadOrder < right.earliestLoadOrder;
            }
            return left.text < right.text;
        });

    result.reserve(std::min(maxCandidates, ranked.size()));

    for (const auto& candidate : ranked) {
        Entry entry;
        entry.text = candidate.text;
        entry.code = candidate.displayCode;
        entry.staticScore = candidate.staticScore;
        entry.loadOrder = candidate.earliestLoadOrder;
        entry.isUser = candidate.hasManualUser;
        entry.isLearned = candidate.hasLearned;
        entry.isAutoPhrase = candidate.hasAutoPhrase;
        result.push_back(std::move(entry));
        if (result.size() >= maxCandidates) {
            break;
        }
    }

    return result;
}

std::vector<std::wstring> CompositionEngine::QueryCandidates(const std::wstring& code, size_t maxCandidates) const {
    std::vector<std::wstring> result;
    const std::vector<Entry> entries = QueryCandidateEntries(code, maxCandidates);
    result.reserve(entries.size());
    for (const auto& item : entries) {
        result.push_back(item.text);
    }
    return result;
}

void CompositionEngine::RecordCommit(const std::wstring& code, const std::wstring& text, std::uint64_t boost) {
    if (code.empty() || text.empty() || boost == 0) {
        return;
    }

    const std::wstring key = MakeFreqKey(NormalizeCode(code), text);
    const auto it = frequency_.find(key);
    if (it == frequency_.end()) {
        frequency_[key] = boost;
    } else {
        it->second += boost;
    }
}
