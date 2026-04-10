#include "CompositionEngine.h"

#include <Windows.h>

#include <algorithm>
#include <filesystem>
#include <fstream>
#include <functional>
#include <sstream>
#include <unordered_map>
#include <utility>

namespace {

constexpr std::uint64_t kFrequencyScoreMax = 255;
constexpr const char* kTextFrequencyCommentPrefix = "#textfreq ";

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
    bool hasSystemFiveCodePhrase = false;
    int shortCodeRank = 100;
    bool shortFullCode = false;
    bool nonGb2312Single = false;
    bool autoOnly = false;
};

bool IsGb2312SingleCandidate(const std::wstring& text) {
    if (text.size() != 1) {
        return false;
    }

    char bytes[4] = {};
    BOOL usedDefaultChar = FALSE;
    const int len = WideCharToMultiByte(
        20936,
        WC_NO_BEST_FIT_CHARS,
        text.c_str(),
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
    return b1 >= 0xB0 && b1 <= 0xF7 && b2 >= 0xA1 && b2 <= 0xFE;
}

bool IsHanText(const std::wstring& text) {
    if (text.empty()) {
        return false;
    }

    for (wchar_t ch : text) {
        const bool isHan =
            (ch >= 0x3400 && ch <= 0x4DBF) ||
            (ch >= 0x4E00 && ch <= 0x9FFF) ||
            (ch >= 0xF900 && ch <= 0xFAFF);
        if (!isHan) {
            return false;
        }
    }

    return true;
}

bool IsFrequencyEligibleEntry(const std::wstring& code, const std::wstring& text) {
    if (text.empty()) {
        return false;
    }

    if (IsGb2312SingleCandidate(text)) {
        return true;
    }

    return text.size() >= 2 && code.size() == 4 && IsHanText(text);
}

std::uint64_t SaturatingAddFrequency(std::uint64_t current, std::uint64_t boost) {
    const std::uint64_t room = kFrequencyScoreMax > current ? (kFrequencyScoreMax - current) : 0;
    return current + std::min(room, boost);
}

bool IsSystemFiveCodePhraseEntry(const CompositionEngine::Entry& entry) {
    return !entry.isUser && entry.text.size() >= 2 && entry.code.size() == 5;
}

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
        score.shortestCodeLength == 2 &&
        score.displayCodeLength == 2) {
        return 1;
    }

    if (score.text.size() == 1 && score.displayCodeLength == 2) {
        return 2;
    }

    if (score.text.size() == 2) {
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

std::wstring BuildPrefixKey(const std::wstring& code, size_t prefixLength) {
    if (code.empty() || prefixLength == 0) {
        return L"";
    }

    return code.substr(0, std::min(prefixLength, code.size()));
}

bool IsBetterSingleCharCodeEntry(const CompositionEngine::Entry& candidate, const CompositionEngine::Entry* bestEntry) {
    if (bestEntry == nullptr) {
        return true;
    }

    // Phrase construction should prefer the longest actual source code and
    // avoid falling back to short-code variants when a full code exists.
    if (candidate.code.size() != bestEntry->code.size()) {
        return candidate.code.size() > bestEntry->code.size();
    }
    if (candidate.staticScore != bestEntry->staticScore) {
        return candidate.staticScore > bestEntry->staticScore;
    }
    return candidate.loadOrder < bestEntry->loadOrder;
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

std::string TrimAsciiWhitespace(const std::string& value) {
    size_t start = 0;
    while (start < value.size() && std::isspace(static_cast<unsigned char>(value[start]))) {
        ++start;
    }

    size_t end = value.size();
    while (end > start && std::isspace(static_cast<unsigned char>(value[end - 1]))) {
        --end;
    }

    return value.substr(start, end - start);
}

std::string ToLowerAscii(std::string value) {
    std::transform(
        value.begin(),
        value.end(),
        value.begin(),
        [](unsigned char ch) {
            return static_cast<char>(std::tolower(ch));
        });
    return value;
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
    if (code.empty()) {
        return false;
    }

    if (code.size() == 1 && part.codeIndex > 0) {
        outCode.push_back(L'v');
        return true;
    }

    // Tolerate shorter source codes by falling back to the last available code letter.
    const size_t safeCodeIndex = std::min(part.codeIndex, code.size() - 1);
    outCode.push_back(code[safeCodeIndex]);
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

std::vector<size_t> BuildMinCodeLengthsForPattern(
    size_t textLength,
    const std::vector<CompositionEngine::PhraseCodePart>& pattern) {
    std::vector<size_t> requiredLengths(textLength, 1);
    for (const CompositionEngine::PhraseCodePart& part : pattern) {
        if (part.noOp) {
            continue;
        }

        const size_t textIndex = part.fromEnd ? (textLength - 1) : part.charIndex;
        if (textIndex >= textLength) {
            continue;
        }

        requiredLengths[textIndex] = std::max(requiredLengths[textIndex], part.codeIndex + 1);
    }

    return requiredLengths;
}

bool IsNonGb2312SingleCandidate(const CandidateScore& score) {
    return score.text.size() == 1 && score.commonCharRank >= 320;
}

void PushUniquePhrasePattern(
    const std::vector<CompositionEngine::PhraseCodePart>& pattern,
    std::vector<std::vector<CompositionEngine::PhraseCodePart>>& outPatterns) {
    if (pattern.empty()) {
        return;
    }

    if (std::find(outPatterns.begin(), outPatterns.end(), pattern) == outPatterns.end()) {
        outPatterns.push_back(pattern);
    }
}

void CollectCompatiblePhrasePatterns(
    size_t textLength,
    const std::vector<CompositionEngine::PhraseCodePart>& primaryPattern,
    std::vector<std::vector<CompositionEngine::PhraseCodePart>>& outPatterns) {
    outPatterns.clear();
    PushUniquePhrasePattern(primaryPattern, outPatterns);

    if (textLength == 2) {
        size_t firstCharTailIndex = 0;
        size_t secondCharTailIndex = 0;
        bool firstCharTailFound = false;
        bool secondCharTailFound = false;
        for (const CompositionEngine::PhraseCodePart& part : primaryPattern) {
            if (part.noOp || part.fromEnd || part.codeIndex == 0) {
                continue;
            }

            if (part.charIndex == 0) {
                firstCharTailIndex = part.codeIndex;
                firstCharTailFound = true;
            } else if (part.charIndex == 1) {
                secondCharTailIndex = part.codeIndex;
                secondCharTailFound = true;
            }
        }

        if (firstCharTailFound && secondCharTailFound) {
            constexpr size_t kMaxTwoCharTailIndex = 3;
            for (size_t firstTailIndex = firstCharTailIndex; firstTailIndex <= kMaxTwoCharTailIndex; ++firstTailIndex) {
                for (size_t secondTailIndex = secondCharTailIndex; secondTailIndex <= kMaxTwoCharTailIndex; ++secondTailIndex) {
                    if (firstTailIndex == firstCharTailIndex && secondTailIndex == secondCharTailIndex) {
                        continue;
                    }

                    std::vector<CompositionEngine::PhraseCodePart> compatiblePattern = primaryPattern;
                    bool adjusted = false;
                    for (CompositionEngine::PhraseCodePart& part : compatiblePattern) {
                        if (part.noOp || part.fromEnd || part.codeIndex == 0) {
                            continue;
                        }

                        if (part.charIndex == 0 && part.codeIndex == firstCharTailIndex) {
                            part.codeIndex = firstTailIndex;
                            adjusted = adjusted || (firstTailIndex != firstCharTailIndex);
                        } else if (part.charIndex == 1 && part.codeIndex == secondCharTailIndex) {
                            part.codeIndex = secondTailIndex;
                            adjusted = adjusted || (secondTailIndex != secondCharTailIndex);
                        }
                    }

                    if (adjusted) {
                        PushUniquePhrasePattern(compatiblePattern, outPatterns);
                    }
                }
            }
        }
        return;
    }

    if (textLength == 3) {
        std::vector<CompositionEngine::PhraseCodePart> secondCharSecondCode = primaryPattern;
        bool adjusted = false;
        for (CompositionEngine::PhraseCodePart& part : secondCharSecondCode) {
            if (!part.noOp && !part.fromEnd && part.charIndex == 1 && part.codeIndex == 2) {
                part.codeIndex = 1;
                adjusted = true;
                break;
            }
        }

        if (adjusted) {
            PushUniquePhrasePattern(secondCharSecondCode, outPatterns);
        }
        return;
    }

    if (textLength >= 5) {
        std::vector<CompositionEngine::PhraseCodePart> lastCharHeadCode = primaryPattern;
        bool adjusted = false;
        for (CompositionEngine::PhraseCodePart& part : lastCharHeadCode) {
            if (!part.noOp && !part.fromEnd && part.charIndex == 3 && part.codeIndex == 0) {
                part.fromEnd = true;
                part.charIndex = 0;
                adjusted = true;
                break;
            }
        }

        if (adjusted) {
            PushUniquePhrasePattern(lastCharHeadCode, outPatterns);
        }
    }
}

void PushUniqueCode(std::vector<std::wstring>& outCodes, const std::wstring& code) {
    if (code.empty()) {
        return;
    }
    if (std::find(outCodes.begin(), outCodes.end(), code) == outCodes.end()) {
        outCodes.push_back(code);
    }
}

bool SelectShortestCodeVariant(
    const std::vector<std::wstring>& variants,
    size_t minLength,
    std::wstring& outCode) {
    outCode.clear();
    for (const std::wstring& variant : variants) {
        if (variant.size() >= minLength) {
            outCode = variant;
            return true;
        }
    }
    return false;
}

bool SelectLongestCodeVariant(
    const std::vector<std::wstring>& variants,
    size_t minLength,
    std::wstring& outCode) {
    outCode.clear();
    for (auto it = variants.rbegin(); it != variants.rend(); ++it) {
        if (it->size() >= minLength) {
            outCode = *it;
            return true;
        }
    }
    return false;
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

CompositionEngine::CandidateKey CompositionEngine::MakeCandidateKey(const std::wstring& code, const std::wstring& text) {
    CandidateKey key;
    key.code = code;
    key.text = text;
    return key;
}

std::wstring CompositionEngine::NormalizeCode(const std::wstring& code) {
    std::wstring normalized;
    normalized.reserve(code.size());

    for (wchar_t ch : code) {
        wchar_t lowered = ch;
        if (lowered >= L'A' && lowered <= L'Z') {
            lowered = static_cast<wchar_t>(lowered - L'A' + L'a');
        }
        if (lowered >= L'a' && lowered <= L'z') {
            normalized.push_back(lowered);
        }
    }

    return normalized;
}

bool CompositionEngine::TryParsePhraseRuleToken(const std::string& token, PhraseCodePart& outPart) {
    outPart = PhraseCodePart{};
    const std::string trimmedToken = TrimAsciiWhitespace(token);
    if (trimmedToken.size() != 3) {
        return false;
    }

    const char direction = static_cast<char>(std::tolower(static_cast<unsigned char>(trimmedToken[0])));
    if (direction != 'p' && direction != 'n') {
        return false;
    }

    if (!std::isdigit(static_cast<unsigned char>(trimmedToken[1])) || !std::isdigit(static_cast<unsigned char>(trimmedToken[2]))) {
        return false;
    }

    const int charOrdinal = trimmedToken[1] - '0';
    const int codeOrdinal = trimmedToken[2] - '0';
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
    const std::string trimmedKey = TrimAsciiWhitespace(key);
    const std::string trimmedValue = TrimAsciiWhitespace(value);
    if (trimmedKey.size() < 2) {
        return false;
    }

    const char scope = static_cast<char>(std::tolower(static_cast<unsigned char>(trimmedKey[0])));
    if (scope != 'e' && scope != 'a') {
        return false;
    }

    size_t length = 0;
    try {
        length = static_cast<size_t>(std::stoul(trimmedKey.substr(1)));
    }
    catch (...) {
        return false;
    }

    if (length == 0) {
        return false;
    }

    std::vector<PhraseCodePart> parts;
    std::istringstream iss(trimmedValue);
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
    const std::string trimmedLine = TrimAsciiWhitespace(line);
    const std::string constructPrefix = "# yuninput:construct_phrase=";
    const std::string rulePrefix = "# yuninput:rule:";
    if (trimmedLine.rfind(constructPrefix, 0) == 0) {
        constructPhrasePrefix_ = Utf8ToWide(trimmedLine.substr(constructPrefix.size()));
        return;
    }

    if (trimmedLine.rfind(rulePrefix, 0) == 0) {
        const std::string spec = trimmedLine.substr(rulePrefix.size());
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

    const Entry* bestEntry = nullptr;
    for (const Entry& entry : entries_) {
        if (entry.text.size() != 1 || entry.text[0] != ch) {
            continue;
        }
        if (entry.code.size() < minLength) {
            continue;
        }
        if (blockedEntries_.find(MakeCandidateKey(entry.code, entry.text)) != blockedEntries_.end()) {
            continue;
        }
        if (IsBetterSingleCharCodeEntry(entry, bestEntry)) {
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

bool CompositionEngine::TryResolvePhrasePattern(size_t textLength, std::vector<PhraseCodePart>& outPattern) const {
    outPattern.clear();
    if (textLength < 2) {
        return false;
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
        {false, 2, 0, false},
        {false, 2, 1, false},
    };
    static const std::vector<PhraseCodePart> kRule4Plus = {
        {false, 0, 0, false},
        {false, 1, 0, false},
        {false, 2, 0, false},
        {false, 3, 0, false},
    };

    const PhraseRule* bestAtLeastRule = nullptr;
    for (const PhraseRule& rule : phraseRules_) {
        if (rule.exactLength) {
            if (rule.length == textLength) {
                outPattern = rule.parts;
                return true;
            }
            continue;
        }

        if (rule.length <= textLength) {
            if (bestAtLeastRule == nullptr || rule.length > bestAtLeastRule->length) {
                bestAtLeastRule = &rule;
            }
        }
    }

    if (bestAtLeastRule != nullptr) {
        outPattern = bestAtLeastRule->parts;
        return true;
    }

    if (textLength == 2) {
        outPattern = kRule2;
        return true;
    }
    if (textLength == 3) {
        outPattern = kRule3;
        return true;
    }

    outPattern = kRule4Plus;
    return true;
}

bool CompositionEngine::TryCollectPhraseCharCodes(
    const std::wstring& text,
    const std::vector<PhraseCodePart>& pattern,
    std::vector<std::wstring>& outCodes) const {
    outCodes.clear();
    if (text.empty()) {
        return false;
    }

    outCodes.reserve(text.size());
    for (size_t i = 0; i < text.size(); ++i) {
        std::wstring charCode;
        if (!TryGetBestSingleCharCode(text[i], 1, charCode)) {
            outCodes.clear();
            return false;
        }
        outCodes.push_back(std::move(charCode));
    }

    return true;
}

bool CompositionEngine::TryBuildPhraseCode(const std::wstring& text, std::wstring& outCode) const {
    outCode.clear();
    if (text.size() < 2) {
        return false;
    }

    std::vector<PhraseCodePart> pattern;
    if (!TryResolvePhrasePattern(text.size(), pattern)) {
        return false;
    }

    std::vector<std::wstring> charCodes;
    if (!TryCollectPhraseCharCodes(text, pattern, charCodes)) {
        return false;
    }

    return BuildPhraseCodeFromPattern(charCodes, pattern, outCode);
}

bool CompositionEngine::TryBuildPhraseCodes(const std::wstring& text, std::vector<std::wstring>& outCodes) const {
    outCodes.clear();
    if (text.size() < 2) {
        return false;
    }

    std::vector<std::vector<std::wstring>> charCodeVariants;
    charCodeVariants.reserve(text.size());
    for (wchar_t ch : text) {
        std::vector<std::wstring> variants;
        if (!TryGetSingleCharCodeVariants(ch, variants) || variants.empty()) {
            if (!outCodes.empty()) {
                return true;
            }
            return false;
        }
        charCodeVariants.push_back(std::move(variants));
    }

    std::vector<PhraseCodePart> pattern;
    if (!TryResolvePhrasePattern(text.size(), pattern)) {
        return !outCodes.empty();
    }

    if (text.size() == 2) {
        std::vector<std::wstring> selectedCodes(2);
        if (!SelectLongestCodeVariant(charCodeVariants[0], 1, selectedCodes[0]) ||
            !SelectLongestCodeVariant(charCodeVariants[1], 1, selectedCodes[1])) {
            return false;
        }

        std::vector<PhraseCodePart> primaryPattern = pattern;
        bool hasPrimaryPattern = false;
        for (PhraseCodePart& part : primaryPattern) {
            if (!part.noOp && !part.fromEnd && part.codeIndex == 1) {
                part.codeIndex = 2;
                hasPrimaryPattern = true;
            }
        }

        std::wstring primaryCode;
        if (hasPrimaryPattern && BuildPhraseCodeFromPattern(selectedCodes, primaryPattern, primaryCode)) {
            PushUniqueCode(outCodes, primaryCode);
        }

        std::wstring compatibleCode;
        if (BuildPhraseCodeFromPattern(selectedCodes, pattern, compatibleCode)) {
            PushUniqueCode(outCodes, compatibleCode);
        }

        return !outCodes.empty();
    }

    if (text.size() == 3) {
        std::vector<std::wstring> selectedCodes(3);
        if (!SelectShortestCodeVariant(charCodeVariants[0], 1, selectedCodes[0]) ||
            !SelectShortestCodeVariant(charCodeVariants[1], 1, selectedCodes[1]) ||
            !SelectShortestCodeVariant(charCodeVariants[2], 1, selectedCodes[2])) {
            return false;
        }

        std::wstring primaryCode;
        if (BuildPhraseCodeFromPattern(selectedCodes, pattern, primaryCode)) {
            PushUniqueCode(outCodes, primaryCode);
        }

        std::vector<PhraseCodePart> compatiblePattern = pattern;
        bool hasCompatiblePattern = false;
        for (PhraseCodePart& part : compatiblePattern) {
            if (!part.noOp && !part.fromEnd && part.charIndex == 1 && part.codeIndex == 2) {
                part.codeIndex = 1;
                hasCompatiblePattern = true;
                break;
            }
        }

        if (hasCompatiblePattern) {
            std::vector<std::wstring> compatibleCodes = selectedCodes;
            std::wstring secondCharCompatibleCode;
            if (SelectLongestCodeVariant(charCodeVariants[1], 2, secondCharCompatibleCode)) {
                compatibleCodes[1] = std::move(secondCharCompatibleCode);
                std::wstring compatibleCode;
                if (BuildPhraseCodeFromPattern(compatibleCodes, compatiblePattern, compatibleCode)) {
                    PushUniqueCode(outCodes, compatibleCode);
                }
            }
        }

        return !outCodes.empty();
    }

    if (text.size() == 4) {
        std::vector<std::wstring> selectedCodes(4);
        for (size_t index = 0; index < selectedCodes.size(); ++index) {
            if (!SelectShortestCodeVariant(charCodeVariants[index], 1, selectedCodes[index])) {
                return false;
            }
        }

        std::wstring primaryCode;
        if (BuildPhraseCodeFromPattern(selectedCodes, pattern, primaryCode)) {
            PushUniqueCode(outCodes, primaryCode);
        }

        return !outCodes.empty();
    }

    if (text.size() >= 5 && text.size() <= 12) {
        std::vector<std::wstring> selectedCodes(text.size());
        for (size_t index = 0; index < selectedCodes.size(); ++index) {
            if (!SelectShortestCodeVariant(charCodeVariants[index], 1, selectedCodes[index])) {
                return false;
            }
        }

        std::wstring primaryCode;
        if (BuildPhraseCodeFromPattern(selectedCodes, pattern, primaryCode)) {
            PushUniqueCode(outCodes, primaryCode);
        }

        std::vector<PhraseCodePart> compatiblePattern = pattern;
        bool hasCompatiblePattern = false;
        for (PhraseCodePart& part : compatiblePattern) {
            if (!part.noOp && !part.fromEnd && part.charIndex == 3 && part.codeIndex == 0) {
                part.fromEnd = true;
                part.charIndex = 0;
                hasCompatiblePattern = true;
                break;
            }
        }

        if (hasCompatiblePattern) {
            std::wstring compatibleCode;
            if (BuildPhraseCodeFromPattern(selectedCodes, compatiblePattern, compatibleCode)) {
                PushUniqueCode(outCodes, compatibleCode);
            }
        }

        return !outCodes.empty();
    }

    std::wstring primaryCode;
    if (TryBuildPhraseCode(text, primaryCode)) {
        PushUniqueCode(outCodes, primaryCode);
    }

    std::vector<std::vector<PhraseCodePart>> compatiblePatterns;
    CollectCompatiblePhrasePatterns(text.size(), pattern, compatiblePatterns);
    if (compatiblePatterns.empty()) {
        compatiblePatterns.push_back(pattern);
    }

    std::wstring building;
    const size_t maxGeneratedCodes = text.size() == 2 ? 256 : (text.size() == 3 ? 128 : 48);
    for (const std::vector<PhraseCodePart>& compatiblePattern : compatiblePatterns) {
        building.clear();
        building.reserve(compatiblePattern.size());

        std::function<void(size_t)> dfs = [&](size_t partIndex) {
            if (outCodes.size() >= maxGeneratedCodes) {
                return;
            }
            if (partIndex >= compatiblePattern.size()) {
                PushUniqueCode(outCodes, building);
                return;
            }

            const PhraseCodePart& part = compatiblePattern[partIndex];
            if (part.noOp) {
                dfs(partIndex + 1);
                return;
            }

            const size_t textIndex = part.fromEnd ? (charCodeVariants.size() - 1) : part.charIndex;
            if (textIndex >= charCodeVariants.size()) {
                return;
            }

            for (const std::wstring& codeVariant : charCodeVariants[textIndex]) {
                if (codeVariant.empty()) {
                    continue;
                }

                const size_t safeCodeIndex = std::min(part.codeIndex, codeVariant.size() - 1);
                building.push_back(codeVariant[safeCodeIndex]);
                dfs(partIndex + 1);
                building.pop_back();
                if (outCodes.size() >= maxGeneratedCodes) {
                    return;
                }
            }
        };

        dfs(0);
        if (outCodes.size() >= maxGeneratedCodes) {
            break;
        }
    }
    return !outCodes.empty();
}

bool CompositionEngine::TryGetSingleCharCodeVariants(wchar_t ch, std::vector<std::wstring>& outCodes) const {
    outCodes.clear();

    std::vector<std::wstring> matchingCodes;
    matchingCodes.reserve(32);
    for (const Entry& entry : entries_) {
        if (entry.text.size() != 1 || entry.text[0] != ch) {
            continue;
        }
        std::wstring normalized = NormalizeCode(entry.code);
        if (normalized.empty()) {
            continue;
        }
        if (blockedEntries_.find(MakeCandidateKey(entry.code, entry.text)) != blockedEntries_.end()) {
            continue;
        }
        if (std::find(matchingCodes.begin(), matchingCodes.end(), normalized) == matchingCodes.end()) {
            matchingCodes.push_back(std::move(normalized));
        }
    }

    std::sort(
        matchingCodes.begin(),
        matchingCodes.end(),
        [](const std::wstring& left, const std::wstring& right) {
            if (left.size() != right.size()) {
                return left.size() < right.size();
            }
            return left < right;
        });

    for (const std::wstring& code : matchingCodes) {
        PushUniqueCode(outCodes, code);
    }

    return !outCodes.empty();
}

int CompositionEngine::GetCommonCharRankCached(const std::wstring& text) const {
    const auto cacheIt = commonCharRankCache_.find(text);
    if (cacheIt != commonCharRankCache_.end()) {
        return cacheIt->second;
    }

    const int rank = GetCommonCharRank(text);
    if (commonCharRankCache_.size() > 8192) {
        commonCharRankCache_.clear();
    }
    commonCharRankCache_.emplace(text, rank);
    return rank;
}

void CompositionEngine::InvalidateQueryCache() const {
    queryCache_.clear();
}

bool CompositionEngine::EntryIndexLess(size_t left, size_t right) const {
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
}

void CompositionEngine::RebuildPrefixRanges() {
    prefixRanges_.clear();
    prefixRanges_.reserve(26 + (26 * 26));
    for (size_t position = 0; position < sortedIndices_.size(); ++position) {
        const Entry& entry = entries_[sortedIndices_[position]];
        if (entry.code.empty()) {
            continue;
        }

        const size_t maxPrefixLength = std::min<size_t>(2, entry.code.size());
        for (size_t prefixLength = 1; prefixLength <= maxPrefixLength; ++prefixLength) {
            const std::wstring prefix = entry.code.substr(0, prefixLength);
            auto it = prefixRanges_.find(prefix);
            if (it == prefixRanges_.end()) {
                prefixRanges_.emplace(prefix, PrefixRange{position, position + 1});
            } else {
                it->second.end = position + 1;
            }
        }
    }
}

void CompositionEngine::InsertEntryIntoIndices(size_t index) {
    const Entry& entry = entries_[index];
    if (entry.isUser) {
        userEntryIndices_.push_back(index);
    }
    if (entry.isAutoPhrase) {
        autoPhraseEntryIndices_.push_back(index);
    }

    const auto insertIt = std::lower_bound(
        sortedIndices_.begin(),
        sortedIndices_.end(),
        index,
        [this](size_t existing, size_t inserted) {
            return EntryIndexLess(existing, inserted);
        });
    sortedIndices_.insert(insertIt, index);
    RebuildPrefixRanges();
    InvalidateQueryCache();
}

void CompositionEngine::RebuildIndex() {
    sortedIndices_.clear();
    sortedIndices_.reserve(entries_.size());
    userEntryIndices_.clear();
    autoPhraseEntryIndices_.clear();
    userEntryIndices_.reserve(entries_.size());
    autoPhraseEntryIndices_.reserve(entries_.size());
    for (size_t i = 0; i < entries_.size(); ++i) {
        sortedIndices_.push_back(i);
        if (entries_[i].isUser) {
            userEntryIndices_.push_back(i);
        }
        if (entries_[i].isAutoPhrase) {
            autoPhraseEntryIndices_.push_back(i);
        }
    }

    std::stable_sort(
        sortedIndices_.begin(),
        sortedIndices_.end(),
        [this](size_t left, size_t right) {
            return EntryIndexLess(left, right);
        });

    RebuildPrefixRanges();

    InvalidateQueryCache();
}

std::pair<std::vector<size_t>::const_iterator, std::vector<size_t>::const_iterator> CompositionEngine::FindCandidateRange(
    const std::wstring& normalizedCode) const {
    if (sortedIndices_.empty()) {
        return {sortedIndices_.end(), sortedIndices_.end()};
    }

    auto rangeBegin = sortedIndices_.begin();
    auto rangeEnd = sortedIndices_.end();
    const std::wstring shortPrefix = BuildPrefixKey(normalizedCode, std::min<size_t>(2, normalizedCode.size()));
    if (!shortPrefix.empty()) {
        const auto prefixIt = prefixRanges_.find(shortPrefix);
        if (prefixIt == prefixRanges_.end()) {
            return {sortedIndices_.end(), sortedIndices_.end()};
        }

        rangeBegin = sortedIndices_.begin() + static_cast<std::ptrdiff_t>(prefixIt->second.begin);
        rangeEnd = sortedIndices_.begin() + static_cast<std::ptrdiff_t>(prefixIt->second.end);
        if (normalizedCode.size() <= 2) {
            return {rangeBegin, rangeEnd};
        }
    }

    const std::wstring rangeEndKey = normalizedCode + std::wstring(1, static_cast<wchar_t>(0xFFFF));
    const auto begin = std::lower_bound(
        rangeBegin,
        rangeEnd,
        normalizedCode,
        [this](size_t idx, const std::wstring& key) {
            return entries_[idx].code < key;
        });

    const auto end = std::lower_bound(
        begin,
        rangeEnd,
        rangeEndKey,
        [this](size_t idx, const std::wstring& key) {
            return entries_[idx].code < key;
        });

    return {begin, end};
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
        bool taggedAuto = false;
        if (iss >> thirdToken) {
            try {
                staticScore = static_cast<std::uint32_t>(std::stoul(thirdToken));
                std::string tagToken;
                while (iss >> tagToken) {
                    if (ToLowerAscii(tagToken) == "auto") {
                        taggedAuto = true;
                    }
                }
            }
            catch (...) {
                if (ToLowerAscii(thirdToken) == "auto") {
                    taggedAuto = true;
                }

                std::string tagToken;
                while (iss >> tagToken) {
                    if (ToLowerAscii(tagToken) == "auto") {
                        taggedAuto = true;
                    }
                }
            }
        }

        Entry entry;
        entry.code = NormalizeCode(Utf8ToWide(codeUtf8));
        entry.text = Utf8ToWide(textUtf8);
        entry.staticScore = staticScore;
        entry.loadOrder = entries_.size();
        entry.isUser = isUserSource;
        entry.isAutoPhrase = isAutoPhraseSource || (isUserSource && taggedAuto);
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

bool CompositionEngine::LoadDictionaryMetadataOnlyFromFile(const std::wstring& filePath) {
    LoadDictionaryMetadataFromFile(filePath + L".rules");
    return true;
}

bool CompositionEngine::LoadFrequencyFromFile(const std::wstring& filePath) {
    frequency_.clear();
    textFrequency_.clear();

    std::ifstream input(filePath, std::ios::in | std::ios::binary);
    if (!input) {
        InvalidateQueryCache();
        return false;
    }

    std::string line;
    while (std::getline(input, line)) {
        if (line.empty()) {
            continue;
        }

        if (line.rfind(kTextFrequencyCommentPrefix, 0) == 0) {
            continue;
        }

        if (line[0] == '#') {
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

        const std::wstring normalizedCode = NormalizeCode(code);
        if (!IsFrequencyEligibleEntry(normalizedCode, text)) {
            continue;
        }

        frequency_[MakeCandidateKey(normalizedCode, text)] = score;
        const auto textIt = textFrequency_.find(text);
        if (textIt == textFrequency_.end() || score > textIt->second) {
            textFrequency_[text] = score;
        }
    }

    InvalidateQueryCache();
    return true;
}

bool CompositionEngine::LoadBlockedEntriesFromFile(const std::wstring& filePath) {
    blockedEntries_.clear();

    std::ifstream input(filePath, std::ios::in | std::ios::binary);
    if (!input) {
        InvalidateQueryCache();
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
        if (!(iss >> codeUtf8 >> textUtf8)) {
            continue;
        }

        const std::wstring code = NormalizeCode(Utf8ToWide(codeUtf8));
        const std::wstring text = Utf8ToWide(textUtf8);
        if (code.empty() || text.empty()) {
            continue;
        }

        blockedEntries_.insert(MakeCandidateKey(code, text));
    }

    InvalidateQueryCache();
    return true;
}

bool CompositionEngine::SaveFrequencyToFile(const std::wstring& filePath) const {
    std::ofstream output(filePath, std::ios::out | std::ios::binary | std::ios::trunc);
    if (!output) {
        return false;
    }

    const std::string content = BuildFrequencyFileContent();
    output.write(content.data(), static_cast<std::streamsize>(content.size()));

    return true;
}

std::string CompositionEngine::BuildFrequencyFileContent() const {
    std::ostringstream output;

    for (const auto& pair : textFrequency_) {
        output << kTextFrequencyCommentPrefix << WideToUtf8(pair.first) << ' ' << pair.second << '\n';
    }

    for (const auto& pair : frequency_) {
        output << WideToUtf8(pair.first.code) << ' ' << WideToUtf8(pair.first.text) << ' ' << pair.second << '\n';
    }

    return output.str();
}

bool CompositionEngine::SaveBlockedEntriesToFile(const std::wstring& filePath) const {
    std::ofstream output(filePath, std::ios::out | std::ios::binary | std::ios::trunc);
    if (!output) {
        return false;
    }

    const std::string content = BuildBlockedEntriesFileContent();
    output.write(content.data(), static_cast<std::streamsize>(content.size()));

    return true;
}

std::string CompositionEngine::BuildBlockedEntriesFileContent() const {
    std::ostringstream output;

    for (const auto& blockedKey : blockedEntries_) {
        output << WideToUtf8(blockedKey.code) << ' ' << WideToUtf8(blockedKey.text) << '\n';
    }

    return output.str();
}

bool CompositionEngine::SaveUserDictionaryToFile(const std::wstring& filePath) const {
    std::ofstream output(filePath, std::ios::out | std::ios::binary | std::ios::trunc);
    if (!output) {
        return false;
    }

    const std::string content = BuildUserDictionaryFileContent();
    output.write(content.data(), static_cast<std::streamsize>(content.size()));

    return true;
}

std::string CompositionEngine::BuildUserDictionaryFileContent() const {
    std::ostringstream output;

    for (size_t index : userEntryIndices_) {
        const Entry& entry = entries_[index];
        if (entry.code.empty() || entry.text.empty()) {
            continue;
        }

        if (entry.isAutoPhrase) {
            continue;
        }

        output << WideToUtf8(entry.code) << ' ' << WideToUtf8(entry.text) << ' ' << entry.staticScore;
        output << '\n';
    }

    return output.str();
}

std::string CompositionEngine::BuildAutoPhraseDictionaryFileContent(size_t maxEntries) const {
    std::ostringstream output;

    size_t startIndex = 0;
    if (maxEntries > 0 && autoPhraseEntryIndices_.size() > maxEntries) {
        startIndex = autoPhraseEntryIndices_.size() - maxEntries;
    }

    for (size_t i = startIndex; i < autoPhraseEntryIndices_.size(); ++i) {
        const Entry& entry = entries_[autoPhraseEntryIndices_[i]];
        if (entry.code.empty() || entry.text.empty()) {
            continue;
        }

        output << WideToUtf8(entry.code) << ' ' << WideToUtf8(entry.text) << ' ' << entry.staticScore << '\n';
    }

    return output.str();
}

bool CompositionEngine::SaveAutoPhraseDictionaryToFile(const std::wstring& filePath) const {
    std::ofstream output(filePath, std::ios::out | std::ios::binary | std::ios::trunc);
    if (!output) {
        return false;
    }

    const std::string content = BuildAutoPhraseDictionaryFileContent();
    output.write(content.data(), static_cast<std::streamsize>(content.size()));

    return true;
}

bool CompositionEngine::AddUserEntry(const std::wstring& code, const std::wstring& text) {
    const std::wstring normalizedCode = NormalizeCode(code);
    if (normalizedCode.empty() || text.empty()) {
        return false;
    }

    blockedEntries_.erase(MakeCandidateKey(normalizedCode, text));

    for (Entry& entry : entries_) {
        if (entry.code == normalizedCode && entry.text == text) {
            const bool changed = !entry.isUser || entry.isAutoPhrase;
            const bool wasAutoPhrase = entry.isAutoPhrase;
            const bool wasUser = entry.isUser;
            entry.isUser = true;
            entry.isAutoPhrase = false;
            if (changed) {
                if (wasUser && wasAutoPhrase) {
                    autoPhraseEntryIndices_.erase(
                        std::remove_if(
                            autoPhraseEntryIndices_.begin(),
                            autoPhraseEntryIndices_.end(),
                            [&](size_t index) {
                                return &entries_[index] == &entry;
                            }),
                        autoPhraseEntryIndices_.end());
                    InvalidateQueryCache();
                } else {
                    RebuildIndex();
                }
            } else {
                InvalidateQueryCache();
            }
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
    InsertEntryIntoIndices(entries_.size() - 1);
    return true;
}

bool CompositionEngine::AddAutoPhraseEntry(const std::wstring& code, const std::wstring& text) {
    const std::wstring normalizedCode = NormalizeCode(code);
    if (normalizedCode.empty() || text.empty()) {
        return false;
    }

    blockedEntries_.erase(MakeCandidateKey(normalizedCode, text));

    for (Entry& entry : entries_) {
        if (entry.code == normalizedCode && entry.text == text) {
            if (entry.isUser && !entry.isAutoPhrase) {
                return false;
            }
            const bool changed = !entry.isAutoPhrase;
            entry.isUser = true;
            entry.isAutoPhrase = true;
            if (changed) {
                RebuildIndex();
            }
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
    InsertEntryIntoIndices(entries_.size() - 1);
    return true;
}

std::unordered_map<wchar_t, std::wstring> CompositionEngine::BuildSingleCharCodeHintMap() const {
    std::unordered_map<wchar_t, const Entry*> bestByChar;
    bestByChar.reserve(4096);

    for (const Entry& entry : entries_) {
        if (entry.text.size() != 1 || entry.code.empty()) {
            continue;
        }
        if (blockedEntries_.find(MakeCandidateKey(entry.code, entry.text)) != blockedEntries_.end()) {
            continue;
        }

        const wchar_t ch = entry.text[0];
        const auto it = bestByChar.find(ch);
        const Entry* currentBest = it == bestByChar.end() ? nullptr : it->second;
        if (IsBetterSingleCharCodeEntry(entry, currentBest)) {
            bestByChar[ch] = &entry;
        }
    }

    std::unordered_map<wchar_t, std::wstring> hints;
    hints.reserve(bestByChar.size());
    for (const auto& pair : bestByChar) {
        hints.emplace(pair.first, pair.second->code);
    }
    return hints;
}

bool CompositionEngine::PinEntry(const std::wstring& code, const std::wstring& text) {
    const std::wstring normalizedCode = NormalizeCode(code);
    if (normalizedCode.empty() || text.empty()) {
        return false;
    }

    blockedEntries_.erase(MakeCandidateKey(normalizedCode, text));

    for (Entry& entry : entries_) {
        if (entry.code == normalizedCode && entry.text == text) {
            const bool changed = !entry.isUser || entry.isAutoPhrase;
            entry.isUser = true;
            entry.isAutoPhrase = false;
            if (changed) {
                RebuildIndex();
            }
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

    const CandidateKey key = MakeCandidateKey(normalizedCode, text);
    const bool inserted = blockedEntries_.insert(key).second;
    frequency_.erase(key);
    if (inserted) {
        InvalidateQueryCache();
    }

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

bool CompositionEngine::HasEntry(const std::wstring& code, const std::wstring& text) const {
    const std::wstring normalizedCode = NormalizeCode(code);
    if (normalizedCode.empty() || text.empty() || sortedIndices_.empty()) {
        return false;
    }

    const auto range = FindCandidateRange(normalizedCode);
    for (auto it = range.first; it != range.second; ++it) {
        const Entry& entry = entries_[*it];
        if (entry.code != normalizedCode || entry.text != text) {
            continue;
        }
        if (blockedEntries_.find(MakeCandidateKey(entry.code, entry.text)) != blockedEntries_.end()) {
            continue;
        }
        return true;
    }

    return false;
}

std::vector<CompositionEngine::Entry> CompositionEngine::QueryCandidateEntries(const std::wstring& code, size_t maxCandidates) const {
    std::vector<Entry> result;
    if (code.empty() || maxCandidates == 0 || sortedIndices_.empty()) {
        return result;
    }

    const std::wstring normalizedCode = NormalizeCode(code);
    const QueryCacheKey cacheKey{normalizedCode, maxCandidates};
    const auto cachedIt = queryCache_.find(cacheKey);
    if (cachedIt != queryCache_.end()) {
        return cachedIt->second;
    }

    const auto range = FindCandidateRange(normalizedCode);
    const auto begin = range.first;
    const auto end = range.second;
    if (begin == sortedIndices_.end() || begin == end) {
        return result;
    }

    result = QueryCandidateEntriesInRange(normalizedCode, begin, end, maxCandidates);

    if (queryCache_.size() >= 256) {
        queryCache_.clear();
    }
    queryCache_.emplace(cacheKey, result);
    return result;
}

std::vector<CompositionEngine::Entry> CompositionEngine::QueryCandidateEntriesFast(
    const std::wstring& code,
    size_t maxCandidates,
    size_t scanBudget) const {
    std::vector<Entry> result;
    if (code.empty() || maxCandidates == 0 || scanBudget == 0 || sortedIndices_.empty()) {
        return result;
    }

    const std::wstring normalizedCode = NormalizeCode(code);
    const auto range = FindCandidateRange(normalizedCode);
    auto begin = range.first;
    auto end = range.second;
    if (begin == sortedIndices_.end() || begin == end) {
        return result;
    }

    const size_t available = static_cast<size_t>(std::distance(begin, end));
    if (available > scanBudget) {
        end = begin + static_cast<std::ptrdiff_t>(scanBudget);
    }

    return QueryCandidateEntriesInRange(normalizedCode, begin, end, maxCandidates);
}

std::vector<CompositionEngine::Entry> CompositionEngine::QueryCandidateEntriesInRange(
    const std::wstring& normalizedCode,
    std::vector<size_t>::const_iterator begin,
    std::vector<size_t>::const_iterator end,
    size_t maxCandidates) const {
    std::vector<Entry> result;
    if (begin == end) {
        return result;
    }

    std::unordered_map<std::wstring, CandidateScore> bestScoreByText;
    bestScoreByText.reserve(static_cast<size_t>(std::distance(begin, end)));

    for (auto it = begin; it != end; ++it) {
        const Entry& item = entries_[*it];
        const CandidateKey freqKey = MakeCandidateKey(item.code, item.text);
        if (blockedEntries_.find(freqKey) != blockedEntries_.end()) {
            continue;
        }
        if (IsLikelyBrokenCandidate(item.text)) {
            continue;
        }
        const bool frequencyEligible = !item.isAutoPhrase && IsFrequencyEligibleEntry(item.code, item.text);
        const auto freqIt = frequency_.find(freqKey);
        const auto textFreqIt = textFrequency_.find(item.text);
        const std::uint64_t codeScore = (frequencyEligible && freqIt != frequency_.end()) ? freqIt->second : 0;
        const std::uint64_t textScore = (frequencyEligible && textFreqIt != textFrequency_.end()) ? textFreqIt->second : 0;
        const std::uint64_t score = std::max(codeScore, textScore);
        const bool exactCode = item.code == normalizedCode;
        const size_t completionDelta = item.code.size() >= normalizedCode.size()
            ? (item.code.size() - normalizedCode.size())
            : 0;
        const int lengthPreferenceScore = GetLengthPreferenceScore(normalizedCode.size(), item.text.size());
        const bool systemFiveCodePhrase = IsSystemFiveCodePhraseEntry(item);

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
            candidate.hasLearned = frequencyEligible && score > 0;
            candidate.commonCharRank = GetCommonCharRankCached(item.text);
            candidate.completionDelta = completionDelta;
            candidate.lengthPreferenceScore = lengthPreferenceScore;
            candidate.preferredTwoCharPhrase = IsPreferredTwoCharPhraseCandidate(candidate, normalizedCode.size());
            candidate.twoCodePriorityTier = GetTwoCodePriorityTier(candidate, normalizedCode.size());
            candidate.hasSystemFiveCodePhrase = systemFiveCodePhrase;
            candidate.shortCodeRank = GetSingleCharShortCodeRank(candidate);
            candidate.shortFullCode = candidate.shortestCodeLength < 4;
            candidate.nonGb2312Single = IsNonGb2312SingleCandidate(candidate);
            candidate.autoOnly = candidate.hasAutoPhrase && !candidate.hasSystemSource && !candidate.hasManualUser;
            bestScoreByText.emplace(item.text, candidate);
            continue;
        }

        CandidateScore& existing = scoreIt->second;
        existing.exactCode = existing.exactCode || exactCode;
        existing.hasUser = existing.hasUser || item.isUser;
        existing.hasManualUser = existing.hasManualUser || (item.isUser && !item.isAutoPhrase);
        existing.hasAutoPhrase = existing.hasAutoPhrase || item.isAutoPhrase;
        existing.hasSystemSource = existing.hasSystemSource || !item.isUser;
        existing.hasLearned = existing.hasLearned || (frequencyEligible && score > 0);
        existing.hasSystemFiveCodePhrase = existing.hasSystemFiveCodePhrase || systemFiveCodePhrase;
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
        const int rank = GetCommonCharRankCached(item.text);
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
        existing.shortCodeRank = GetSingleCharShortCodeRank(existing);
        existing.shortFullCode = existing.shortestCodeLength < 4;
        existing.nonGb2312Single = IsNonGb2312SingleCandidate(existing);
        existing.autoOnly = existing.hasAutoPhrase && !existing.hasSystemSource && !existing.hasManualUser;

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

    const auto candidateLess = [](const auto& left, const auto& right) {
            if (left.exactCode != right.exactCode) {
                return left.exactCode > right.exactCode;
            }
            if (left.hasManualUser != right.hasManualUser) {
                return left.hasManualUser > right.hasManualUser;
            }
            if (left.hasLearned != right.hasLearned) {
                return left.hasLearned > right.hasLearned;
            }
            if (left.frequency != right.frequency) {
                return left.frequency > right.frequency;
            }
            if (left.twoCodePriorityTier != right.twoCodePriorityTier) {
                return left.twoCodePriorityTier < right.twoCodePriorityTier;
            }
            if (left.preferredTwoCharPhrase != right.preferredTwoCharPhrase) {
                return left.preferredTwoCharPhrase > right.preferredTwoCharPhrase;
            }
            if (left.shortFullCode != right.shortFullCode) {
                return left.shortFullCode > right.shortFullCode;
            }
            if (left.shortCodeRank != right.shortCodeRank) {
                return left.shortCodeRank < right.shortCodeRank;
            }
            if (left.commonCharRank != right.commonCharRank) {
                return left.commonCharRank < right.commonCharRank;
            }
            if (left.hasSystemFiveCodePhrase != right.hasSystemFiveCodePhrase) {
                return left.hasSystemFiveCodePhrase < right.hasSystemFiveCodePhrase;
            }
            if (left.nonGb2312Single != right.nonGb2312Single) {
                return left.nonGb2312Single < right.nonGb2312Single;
            }
            if (left.autoOnly != right.autoOnly) {
                return left.autoOnly < right.autoOnly;
            }
            if (left.shortestCodeLength != right.shortestCodeLength) {
                // Shorter code level should come first when user/learned signals tie.
                return left.shortestCodeLength < right.shortestCodeLength;
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
        };

    const size_t rankedLimit = std::min(maxCandidates, ranked.size());
    if (rankedLimit < ranked.size()) {
        std::partial_sort(ranked.begin(), ranked.begin() + rankedLimit, ranked.end(), candidateLess);
        ranked.resize(rankedLimit);
    } else {
        std::stable_sort(ranked.begin(), ranked.end(), candidateLess);
    }

    result.reserve(ranked.size());

    for (const auto& candidate : ranked) {
        Entry entry;
        entry.text = candidate.text;
        entry.code = candidate.displayCode;
        entry.staticScore = candidate.staticScore;
        entry.learnedScore = candidate.frequency;
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

    const std::wstring normalizedCode = NormalizeCode(code);
    if (!IsFrequencyEligibleEntry(normalizedCode, text)) {
        return;
    }

    const CandidateKey key = MakeCandidateKey(normalizedCode, text);
    std::uint64_t current = 0;
    const auto it = frequency_.find(key);
    if (it != frequency_.end()) {
        current = it->second;
    }

    frequency_[key] = SaturatingAddFrequency(current, boost);

    std::uint64_t textCurrent = 0;
    const auto textIt = textFrequency_.find(text);
    if (textIt != textFrequency_.end()) {
        textCurrent = textIt->second;
    }
    textFrequency_[text] = SaturatingAddFrequency(textCurrent, boost);
    InvalidateQueryCache();
}
