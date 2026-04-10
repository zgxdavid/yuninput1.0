#include "CompositionEngine.h"

#include <Windows.h>

#include <algorithm>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <set>
#include <sstream>
#include <string>
#include <thread>
#include <unordered_map>
#include <unordered_set>
#include <vector>

namespace {

struct DictEntry {
    std::wstring code;
    std::wstring text;
    size_t sourcePriority = 0;
    size_t sourceOrder = 0;
};

struct SessionSnapshot {
    ULONGLONG bootTick = 0;
    bool hasBootTick = false;
    std::wstring history;
};

constexpr wchar_t kSessionAutoPhraseBreak = static_cast<wchar_t>(0xE000);
constexpr size_t kSessionAutoPhraseMaxLength = 12;
constexpr const wchar_t* kAutoPhraseHelperMutexName = L"Local\\Yuninput.AutoPhraseHelper.Singleton";
constexpr const wchar_t* kAutoPhraseHelperWakeEventName = L"Local\\Yuninput.AutoPhraseHelper.Wake";
constexpr DWORD kSessionWatchPollIntervalMs = 1500U;

std::wstring Utf8ToWide(const std::string& input) {
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

    decoded = decode(54936, 0);
    if (!decoded.empty()) {
        return decoded;
    }

    return decode(CP_ACP, 0);
}

std::string WideToUtf8(const std::wstring& input) {
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

bool IsCodeToken(const std::string& token) {
    if (token.empty()) {
        return false;
    }

    for (unsigned char ch : token) {
        if (!((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z'))) {
            return false;
        }
    }

    return true;
}

std::wstring NormalizeCode(const std::wstring& code) {
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

bool IsHanCharacter(wchar_t ch) {
    return (ch >= 0x3400 && ch <= 0x4DBF) ||
           (ch >= 0x4E00 && ch <= 0x9FFF) ||
           (ch >= 0xF900 && ch <= 0xFAFF);
}

std::wstring NormalizeSessionPhraseCode(std::wstring code) {
    code.erase(
        std::remove_if(
            code.begin(),
            code.end(),
            [](wchar_t ch) {
                return ch < L'a' || ch > L'z';
            }),
        code.end());
    if (code.size() > 20) {
        code.resize(20);
    }
    return code;
}

std::wstring MakeEntryKey(const std::wstring& code, const std::wstring& text) {
    return NormalizeCode(code) + L"\t" + text;
}

bool TryParseDictionaryEntry(const std::string& line, std::wstring& outCode, std::wstring& outText) {
    outCode.clear();
    outText.clear();
    if (line.empty() || line[0] == '#') {
        return false;
    }

    std::istringstream iss(line);
    std::string firstToken;
    std::string secondToken;
    if (!(iss >> firstToken >> secondToken)) {
        return false;
    }

    const bool firstIsCode = IsCodeToken(firstToken);
    const bool secondIsCode = IsCodeToken(secondToken);
    if (firstIsCode == secondIsCode) {
        return false;
    }

    const std::string codeUtf8 = firstIsCode ? firstToken : secondToken;
    const std::string textUtf8 = firstIsCode ? secondToken : firstToken;
    outCode = NormalizeCode(Utf8ToWide(codeUtf8));
    outText = Utf8ToWide(textUtf8);
    return !outCode.empty() && !outText.empty();
}

bool IsGb2312SingleChar(wchar_t ch) {
    char buffer[4] = {};
    BOOL usedDefault = FALSE;
    const int encoded = WideCharToMultiByte(936, WC_NO_BEST_FIT_CHARS, &ch, 1, buffer, sizeof(buffer), nullptr, &usedDefault);
    if (encoded <= 0 || usedDefault) {
        return false;
    }

    wchar_t roundTrip = 0;
    const int decoded = MultiByteToWideChar(936, MB_ERR_INVALID_CHARS, buffer, encoded, &roundTrip, 1);
    return decoded == 1 && roundTrip == ch;
}

int GetSortTier(const DictEntry& entry) {
    const bool singleChar = entry.text.size() == 1;
    const size_t codeLength = entry.code.size();
    if (singleChar && !IsGb2312SingleChar(entry.text[0])) {
        return 7;
    }
    if (singleChar) {
        if (codeLength <= 1) {
            return 0;
        }
        if (codeLength == 2) {
            return 2;
        }
        if (codeLength == 3) {
            return 3;
        }
        return 4;
    }

    if (entry.text.size() == 2 && codeLength == 2) {
        return 1;
    }
    if (codeLength >= 5) {
        return 6;
    }
    return 5;
}

bool EnsureDictionaryFileExists(const std::wstring& filePath, const std::string& title) {
    namespace fs = std::filesystem;

    std::error_code ec;
    const fs::path path(filePath);
    if (fs::exists(path, ec)) {
        return true;
    }
    if (path.has_parent_path()) {
        fs::create_directories(path.parent_path(), ec);
    }

    std::ofstream output(path, std::ios::out | std::ios::binary | std::ios::trunc);
    if (!output) {
        return false;
    }
    output << "# " << title << "\n";
    output << "# format: code text\n";
    return true;
}

std::wstring QueryFileStampToken(const std::wstring& filePath) {
    if (filePath.empty()) {
        return L"-";
    }

    std::error_code ec;
    const std::filesystem::path path(filePath);
    const bool exists = std::filesystem::exists(path, ec);
    if (ec || !exists) {
        return L"0";
    }

    ec.clear();
    const auto writeTime = std::filesystem::last_write_time(path, ec);
    if (ec) {
        return L"1";
    }

    return L"1:" + std::to_wstring(static_cast<long long>(writeTime.time_since_epoch().count()));
}

bool AppendAutoPhraseEntries(
    const std::wstring& filePath,
    const std::vector<std::pair<std::wstring, std::wstring>>& additions) {
    if (filePath.empty()) {
        return false;
    }
    if (additions.empty()) {
        return true;
    }

    std::error_code ec;
    const std::filesystem::path path(filePath);
    const bool existed = std::filesystem::exists(path, ec);
    if (path.has_parent_path()) {
        std::filesystem::create_directories(path.parent_path(), ec);
    }

    std::ofstream output(path, std::ios::out | std::ios::binary | std::ios::app);
    if (!output) {
        return false;
    }

    if (!existed) {
        output << "# helper runtime auto phrase dictionary\n";
        output << "# format: code text score\n";
    }

    for (const auto& addition : additions) {
        output << WideToUtf8(addition.first) << ' ' << WideToUtf8(addition.second) << " 1\n";
    }

    return output.good();
}

bool LoadBaseDictionary(CompositionEngine& engine, const std::wstring& primaryDictPath, const std::wstring& fallbackDictPath) {
    bool loadedPrimary = false;
    if (!primaryDictPath.empty() && std::filesystem::exists(primaryDictPath)) {
        loadedPrimary = engine.LoadDictionaryFromFile(primaryDictPath);
    }
    if (!loadedPrimary && !fallbackDictPath.empty() && std::filesystem::exists(fallbackDictPath)) {
        loadedPrimary = engine.LoadDictionaryFromFile(fallbackDictPath);
    }
    return loadedPrimary;
}

bool LoadPhraseSourceDictionary(
    CompositionEngine& engine,
    const std::wstring& phraseSourceDictPath,
    const std::wstring& primaryDictPath,
    const std::wstring& fallbackDictPath) {
    bool loadedPhraseSource = false;
    if (!phraseSourceDictPath.empty() && std::filesystem::exists(phraseSourceDictPath)) {
        loadedPhraseSource = engine.LoadDictionaryFromFile(phraseSourceDictPath);
    }
    if (!loadedPhraseSource && !primaryDictPath.empty() && std::filesystem::exists(primaryDictPath)) {
        loadedPhraseSource = engine.LoadDictionaryFromFile(primaryDictPath);
    }
    if (!loadedPhraseSource && !fallbackDictPath.empty() && std::filesystem::exists(fallbackDictPath)) {
        loadedPhraseSource = engine.LoadDictionaryFromFile(fallbackDictPath);
    }

    if (loadedPhraseSource) {
        if (!phraseSourceDictPath.empty()) {
            engine.LoadDictionaryMetadataOnlyFromFile(phraseSourceDictPath);
        }
        if (!primaryDictPath.empty()) {
            engine.LoadDictionaryMetadataOnlyFromFile(primaryDictPath);
        }
        if (!fallbackDictPath.empty()) {
            engine.LoadDictionaryMetadataOnlyFromFile(fallbackDictPath);
        }
    }

    return loadedPhraseSource;
}

bool CollectSessionAutoPhraseAdditions(
    const SessionSnapshot& snapshot,
    const std::wstring& primaryDictPath,
    const std::wstring& fallbackDictPath,
    const std::wstring& phraseSourceDictPath,
    const std::wstring& userDictPath,
    const std::wstring& extendDictPath,
    const std::vector<std::wstring>& runtimeDictPaths,
    std::vector<std::pair<std::wstring, std::wstring>>& outAdditions) {
    outAdditions.clear();

    CompositionEngine dedupeEngine;
    if (!LoadBaseDictionary(dedupeEngine, primaryDictPath, fallbackDictPath)) {
        return false;
    }
    if (!userDictPath.empty() && std::filesystem::exists(userDictPath)) {
        dedupeEngine.LoadUserDictionaryFromFile(userDictPath);
    }
    if (!extendDictPath.empty() && std::filesystem::exists(extendDictPath)) {
        dedupeEngine.LoadAutoPhraseDictionaryFromFile(extendDictPath);
    }
    for (const std::wstring& runtimeDictPath : runtimeDictPaths) {
        if (!runtimeDictPath.empty() && std::filesystem::exists(runtimeDictPath)) {
            dedupeEngine.LoadAutoPhraseDictionaryFromFile(runtimeDictPath);
        }
    }

    CompositionEngine phraseEngine;
    if (!LoadPhraseSourceDictionary(phraseEngine, phraseSourceDictPath, primaryDictPath, fallbackDictPath)) {
        return false;
    }

    std::unordered_map<std::wstring, std::vector<std::wstring>> candidateCodes;
    for (size_t segmentStart = 0; segmentStart < snapshot.history.size();) {
        while (segmentStart < snapshot.history.size() && !IsHanCharacter(snapshot.history[segmentStart])) {
            ++segmentStart;
        }
        if (segmentStart >= snapshot.history.size()) {
            break;
        }

        size_t segmentEnd = segmentStart;
        while (segmentEnd < snapshot.history.size() && IsHanCharacter(snapshot.history[segmentEnd])) {
            ++segmentEnd;
        }

        const size_t segmentLength = segmentEnd - segmentStart;
        if (segmentLength >= 2) {
            const size_t maxPhraseLength = std::min(kSessionAutoPhraseMaxLength, segmentLength);
            for (size_t start = 0; start + 2 <= segmentLength; ++start) {
                const size_t remaining = segmentLength - start;
                const size_t maxLengthForStart = std::min(maxPhraseLength, remaining);
                for (size_t phraseLength = 2; phraseLength <= maxLengthForStart; ++phraseLength) {
                    const std::wstring phraseText = snapshot.history.substr(segmentStart + start, phraseLength);
                    std::vector<std::wstring> phraseCodes;
                    if (!phraseEngine.TryBuildPhraseCodes(phraseText, phraseCodes) || phraseCodes.empty()) {
                        continue;
                    }

                    auto& codes = candidateCodes[phraseText];
                    for (std::wstring phraseCode : phraseCodes) {
                        phraseCode = NormalizeSessionPhraseCode(std::move(phraseCode));
                        if (phraseCode.empty() || dedupeEngine.HasEntry(phraseCode, phraseText)) {
                            continue;
                        }
                        if (std::find(codes.begin(), codes.end(), phraseCode) != codes.end()) {
                            continue;
                        }
                        codes.push_back(std::move(phraseCode));
                    }

                    if (codes.empty()) {
                        candidateCodes.erase(phraseText);
                    }
                }
            }
        }

        segmentStart = segmentEnd + 1;
    }

    for (auto& pair : candidateCodes) {
        for (std::wstring& code : pair.second) {
            outAdditions.emplace_back(code, pair.first);
        }
    }

    std::sort(
        outAdditions.begin(),
        outAdditions.end(),
        [](const auto& left, const auto& right) {
            if (left.second.size() != right.second.size()) {
                return left.second.size() > right.second.size();
            }
            if (left.second != right.second) {
                return left.second < right.second;
            }
            return left.first < right.first;
        });

    return true;
}

bool ParseDictionaryFile(
    const std::wstring& filePath,
    size_t sourcePriority,
    std::vector<DictEntry>* outEntries,
    std::unordered_set<std::wstring>* outPairs,
    std::set<std::wstring>* outPhraseTexts,
    std::vector<std::string>* outRawLines = nullptr) {
    std::ifstream input(std::filesystem::path(filePath), std::ios::in | std::ios::binary);
    if (!input) {
        return false;
    }

    std::string line;
    size_t sourceOrder = 0;
    while (std::getline(input, line)) {
        if (outRawLines != nullptr) {
            outRawLines->push_back(line);
        }

        std::wstring code;
        std::wstring text;
        if (!TryParseDictionaryEntry(line, code, text)) {
            continue;
        }

        const std::wstring key = MakeEntryKey(code, text);
        if (outPairs != nullptr && !outPairs->insert(key).second) {
            ++sourceOrder;
            continue;
        }

        if (outEntries != nullptr) {
            outEntries->push_back(DictEntry{code, text, sourcePriority, sourceOrder});
        }
        if (outPhraseTexts != nullptr && text.size() >= 2 && text.size() <= 64) {
            outPhraseTexts->insert(text);
        }
        ++sourceOrder;
    }

    return true;
}

bool WriteDictionaryFile(const std::wstring& filePath, const std::vector<DictEntry>& entries, bool singleOnly) {
    std::ofstream output(std::filesystem::path(filePath), std::ios::out | std::ios::binary | std::ios::trunc);
    if (!output) {
        return false;
    }

    output << "# generated dictionary\n";
    output << "# format: code text\n";
    for (const DictEntry& entry : entries) {
        if (singleOnly && entry.text.size() != 1) {
            continue;
        }
        output << WideToUtf8(entry.code) << ' ' << WideToUtf8(entry.text) << '\n';
    }
    return true;
}

bool LoadSessionSnapshot(const std::wstring& filePath, SessionSnapshot* outSnapshot) {
    if (outSnapshot == nullptr) {
        return false;
    }

    std::ifstream input(std::filesystem::path(filePath), std::ios::in | std::ios::binary);
    if (!input) {
        return false;
    }

    *outSnapshot = SessionSnapshot{};
    std::string line;
    while (std::getline(input, line)) {
        if (line.empty()) {
            continue;
        }

        std::istringstream iss(line);
        std::string tag;
        if (!std::getline(iss, tag, '\t')) {
            continue;
        }

        if (tag == "boot_tick") {
            std::string tickText;
            if (std::getline(iss, tickText, '\t')) {
                outSnapshot->bootTick = static_cast<ULONGLONG>(_strtoui64(tickText.c_str(), nullptr, 10));
                outSnapshot->hasBootTick = true;
            }
            continue;
        }

        if (tag == "history") {
            std::string historyUtf8;
            if (std::getline(iss, historyUtf8, '\t')) {
                outSnapshot->history = Utf8ToWide(historyUtf8);
            }
        }
    }

    return true;
}

bool WriteSessionSnapshot(
    const std::wstring& filePath,
    const SessionSnapshot& snapshot,
    const std::vector<std::pair<std::wstring, std::vector<std::wstring>>>& entries) {
    const std::filesystem::path path(filePath);
    std::error_code ec;
    if (path.has_parent_path()) {
        std::filesystem::create_directories(path.parent_path(), ec);
        ec.clear();
    }

    const std::filesystem::path tempPath = path.parent_path() / (path.filename().wstring() + L".tmp");
    std::ofstream output(tempPath, std::ios::out | std::ios::binary | std::ios::trunc);
    if (!output) {
        return false;
    }

    if (snapshot.hasBootTick) {
        output << "boot_tick\t" << snapshot.bootTick << '\n';
    }
    output << "history\t" << WideToUtf8(snapshot.history) << '\n';
    for (const auto& entry : entries) {
        if (entry.first.empty() || entry.second.empty()) {
            continue;
        }

        output << "entry\t" << WideToUtf8(entry.first) << '\t';
        for (size_t index = 0; index < entry.second.size(); ++index) {
            if (index != 0) {
                output << ',';
            }
            output << WideToUtf8(entry.second[index]);
        }
        output << "\t0\n";
    }
    output.close();
    if (!output) {
        return false;
    }

    std::filesystem::remove(path, ec);
    ec.clear();
    std::filesystem::rename(tempPath, path, ec);
    if (ec) {
        std::filesystem::remove(tempPath, ec);
        return false;
    }

    return true;
}

int RunMergeCommand(int argc, char* argv[]) {
    if (argc < 6) {
        std::cerr << "usage: yuninput_user_dict_builder merge <all-dict> <single-dict> <source-dict-1> <source-dict-2> [more-source-dicts...]\n";
        return 1;
    }

    const std::wstring allDictPath = Utf8ToWide(argv[2]);
    const std::wstring singleDictPath = Utf8ToWide(argv[3]);
    std::vector<DictEntry> entries;
    entries.reserve(262144);
    std::unordered_set<std::wstring> seenPairs;
    seenPairs.reserve(262144);

    for (int i = 4; i < argc; ++i) {
        if (!ParseDictionaryFile(Utf8ToWide(argv[i]), static_cast<size_t>(i - 4), &entries, &seenPairs, nullptr)) {
            std::cerr << "failed to parse source dictionary\n";
            return 2;
        }
    }

    std::stable_sort(
        entries.begin(),
        entries.end(),
        [](const DictEntry& left, const DictEntry& right) {
            const int leftTier = GetSortTier(left);
            const int rightTier = GetSortTier(right);
            if (leftTier != rightTier) {
                return leftTier < rightTier;
            }
            if (left.sourcePriority != right.sourcePriority) {
                return left.sourcePriority < right.sourcePriority;
            }
            return left.sourceOrder < right.sourceOrder;
        });

    if (!WriteDictionaryFile(allDictPath, entries, false)) {
        std::cerr << "failed to write merged dictionary\n";
        return 3;
    }
    if (!WriteDictionaryFile(singleDictPath, entries, true)) {
        std::cerr << "failed to write single-char dictionary\n";
        return 4;
    }

    std::cout << "entries_total=" << entries.size() << '\n';
    std::cout << "output_all=" << WideToUtf8(allDictPath) << '\n';
    std::cout << "output_single=" << WideToUtf8(singleDictPath) << '\n';
    return 0;
}

int RunExtendCommand(int argc, char* argv[]) {
    if (argc < 6) {
        std::cerr << "usage: yuninput_user_dict_builder extend <extend-dict> <rules-dict> <single-dict> <phrase-dict> [dedupe-dict...]\n";
        return 1;
    }

    const std::wstring extendDictPath = Utf8ToWide(argv[2]);
    const std::wstring rulesDictPath = Utf8ToWide(argv[3]);
    const std::wstring singleDictPath = Utf8ToWide(argv[4]);
    const std::wstring phraseDictPath = Utf8ToWide(argv[5]);

    if (!EnsureDictionaryFileExists(extendDictPath, "user extend dictionary")) {
        std::cerr << "failed to initialize extend dictionary\n";
        return 2;
    }

    std::unordered_set<std::wstring> existingPairs;
    existingPairs.reserve(262144);
    std::set<std::wstring> phraseTexts;
    std::vector<std::string> extendRawLines;
    if (!ParseDictionaryFile(phraseDictPath, 0, nullptr, &existingPairs, &phraseTexts)) {
        std::cerr << "failed to parse phrase dictionary\n";
        return 3;
    }
    if (!ParseDictionaryFile(extendDictPath, 0, nullptr, &existingPairs, nullptr, &extendRawLines)) {
        std::cerr << "failed to parse extend dictionary\n";
        return 4;
    }
    for (int i = 6; i < argc; ++i) {
        if (!ParseDictionaryFile(Utf8ToWide(argv[i]), 0, nullptr, &existingPairs, nullptr)) {
            std::cerr << "failed to parse dedupe dictionary\n";
            return 5;
        }
    }

    CompositionEngine engine;
    if (!engine.LoadDictionaryFromFile(singleDictPath)) {
        std::cerr << "failed to load single dictionary\n";
        return 6;
    }
    if (!engine.LoadUserDictionaryFromFile(rulesDictPath)) {
        std::cerr << "failed to load rules dictionary\n";
        return 7;
    }
    engine.LoadAutoPhraseDictionaryFromFile(extendDictPath);

    std::vector<std::pair<std::wstring, std::wstring>> additions;
    additions.reserve(8192);
    for (const std::wstring& phraseText : phraseTexts) {
        std::vector<std::wstring> phraseCodes;
        if (!engine.TryBuildPhraseCodes(phraseText, phraseCodes) || phraseCodes.empty()) {
            continue;
        }

        for (const std::wstring& phraseCode : phraseCodes) {
            const std::wstring key = MakeEntryKey(phraseCode, phraseText);
            if (!existingPairs.insert(key).second) {
                continue;
            }
            additions.emplace_back(NormalizeCode(phraseCode), phraseText);
        }
    }

    std::sort(
        additions.begin(),
        additions.end(),
        [](const auto& left, const auto& right) {
            if (left.second != right.second) {
                return left.second < right.second;
            }
            return left.first < right.first;
        });

    if (!additions.empty()) {
        std::ofstream output(std::filesystem::path(extendDictPath), std::ios::out | std::ios::binary | std::ios::trunc);
        if (!output) {
            std::cerr << "failed to rewrite extend dictionary\n";
            return 8;
        }

        if (extendRawLines.empty()) {
            output << "# user extend dictionary\n";
            output << "# format: code text\n";
        } else {
            for (const std::string& line : extendRawLines) {
                output << line << '\n';
            }
        }

        for (const auto& addition : additions) {
            output << WideToUtf8(addition.first) << ' ' << WideToUtf8(addition.second) << '\n';
        }
    }

    std::cout << "phrases_scanned=" << phraseTexts.size() << '\n';
    std::cout << "codes_added=" << additions.size() << '\n';
    std::cout << "output=" << WideToUtf8(extendDictPath) << '\n';
    return 0;
}

int RunSessionBuildCommand(int argc, char* argv[]) {
    if (argc < 8) {
        std::cerr << "usage: yuninput_user_dict_builder session-build <session-file> <primary-dict> <fallback-dict> <phrase-source-dict> <user-dict> <extend-dict>\n";
        return 1;
    }

    const std::wstring sessionFilePath = Utf8ToWide(argv[2]);
    const std::wstring primaryDictPath = Utf8ToWide(argv[3]);
    const std::wstring fallbackDictPath = Utf8ToWide(argv[4]);
    const std::wstring phraseSourceDictPath = Utf8ToWide(argv[5]);
    const std::wstring userDictPath = Utf8ToWide(argv[6]);
    const std::wstring extendDictPath = Utf8ToWide(argv[7]);

    SessionSnapshot snapshot;
    if (!LoadSessionSnapshot(sessionFilePath, &snapshot)) {
        std::cerr << "failed to load session snapshot\n";
        return 2;
    }

    CompositionEngine dedupeEngine;
    if (!LoadBaseDictionary(dedupeEngine, primaryDictPath, fallbackDictPath)) {
        std::cerr << "failed to load base dictionary\n";
        return 3;
    }
    if (!userDictPath.empty() && std::filesystem::exists(userDictPath)) {
        dedupeEngine.LoadUserDictionaryFromFile(userDictPath);
    }
    if (!extendDictPath.empty() && std::filesystem::exists(extendDictPath)) {
        dedupeEngine.LoadAutoPhraseDictionaryFromFile(extendDictPath);
    }

    CompositionEngine phraseEngine;
    if (!LoadPhraseSourceDictionary(phraseEngine, phraseSourceDictPath, primaryDictPath, fallbackDictPath)) {
        std::cerr << "failed to load phrase source dictionary\n";
        return 4;
    }

    std::unordered_map<std::wstring, std::vector<std::wstring>> candidateCodes;
    for (size_t segmentStart = 0; segmentStart < snapshot.history.size();) {
        while (segmentStart < snapshot.history.size() && !IsHanCharacter(snapshot.history[segmentStart])) {
            ++segmentStart;
        }
        if (segmentStart >= snapshot.history.size()) {
            break;
        }

        size_t segmentEnd = segmentStart;
        while (segmentEnd < snapshot.history.size() && IsHanCharacter(snapshot.history[segmentEnd])) {
            ++segmentEnd;
        }

        const size_t segmentLength = segmentEnd - segmentStart;
        if (segmentLength >= 2) {
            const size_t maxPhraseLength = std::min(kSessionAutoPhraseMaxLength, segmentLength);
            for (size_t start = 0; start + 2 <= segmentLength; ++start) {
                const size_t remaining = segmentLength - start;
                const size_t maxLengthForStart = std::min(maxPhraseLength, remaining);
                for (size_t phraseLength = 2; phraseLength <= maxLengthForStart; ++phraseLength) {
                    const std::wstring phraseText = snapshot.history.substr(segmentStart + start, phraseLength);
                    std::vector<std::wstring> phraseCodes;
                    if (!phraseEngine.TryBuildPhraseCodes(phraseText, phraseCodes) || phraseCodes.empty()) {
                        continue;
                    }

                    auto& codes = candidateCodes[phraseText];
                    for (std::wstring phraseCode : phraseCodes) {
                        phraseCode = NormalizeSessionPhraseCode(std::move(phraseCode));
                        if (phraseCode.empty() || dedupeEngine.HasEntry(phraseCode, phraseText)) {
                            continue;
                        }
                        if (std::find(codes.begin(), codes.end(), phraseCode) != codes.end()) {
                            continue;
                        }
                        codes.push_back(std::move(phraseCode));
                    }

                    if (codes.empty()) {
                        candidateCodes.erase(phraseText);
                    }
                }
            }
        }

        segmentStart = segmentEnd + 1;
    }

    std::vector<std::pair<std::wstring, std::vector<std::wstring>>> entries;
    entries.reserve(candidateCodes.size());
    for (auto& pair : candidateCodes) {
        if (!pair.first.empty() && !pair.second.empty()) {
            std::sort(pair.second.begin(), pair.second.end());
            entries.emplace_back(pair.first, std::move(pair.second));
        }
    }
    std::sort(
        entries.begin(),
        entries.end(),
        [](const auto& left, const auto& right) {
            if (left.first.size() != right.first.size()) {
                return left.first.size() > right.first.size();
            }
            return left.first < right.first;
        });

    if (!WriteSessionSnapshot(sessionFilePath, snapshot, entries)) {
        std::cerr << "failed to write session snapshot\n";
        return 5;
    }

    std::cout << "phrases_built=" << entries.size() << '\n';
    std::cout << "output=" << WideToUtf8(sessionFilePath) << '\n';
    return 0;
}

int RunSessionWatchCommand(int argc, char* argv[]) {
    if (argc < 10) {
        std::cerr << "usage: yuninput_user_dict_builder session-watch <session-file> <primary-dict> <fallback-dict> <phrase-source-dict> <user-dict> <extend-dict> <runtime-dict> <helper-dict>\n";
        return 1;
    }

    const std::wstring sessionFilePath = Utf8ToWide(argv[2]);
    const std::wstring primaryDictPath = Utf8ToWide(argv[3]);
    const std::wstring fallbackDictPath = Utf8ToWide(argv[4]);
    const std::wstring phraseSourceDictPath = Utf8ToWide(argv[5]);
    const std::wstring userDictPath = Utf8ToWide(argv[6]);
    const std::wstring extendDictPath = Utf8ToWide(argv[7]);
    const std::wstring runtimeDictPath = Utf8ToWide(argv[8]);
    const std::wstring helperDictPath = Utf8ToWide(argv[9]);

    HANDLE helperMutex = CreateMutexW(nullptr, FALSE, kAutoPhraseHelperMutexName);
    if (helperMutex == nullptr) {
        std::cerr << "failed to create helper mutex\n";
        return 2;
    }
    const bool helperAlreadyRunning = GetLastError() == ERROR_ALREADY_EXISTS;

    HANDLE wakeEvent = CreateEventW(nullptr, FALSE, FALSE, kAutoPhraseHelperWakeEventName);
    if (wakeEvent == nullptr) {
        CloseHandle(helperMutex);
        std::cerr << "failed to create helper wake event\n";
        return 3;
    }

    if (helperAlreadyRunning) {
        SetEvent(wakeEvent);
        CloseHandle(wakeEvent);
        CloseHandle(helperMutex);
        return 0;
    }

    std::wstring lastSessionStamp;
    for (;;) {
        const std::wstring sessionStamp = QueryFileStampToken(sessionFilePath);
        if (sessionStamp != lastSessionStamp) {
            lastSessionStamp = sessionStamp;

            SessionSnapshot snapshot;
            if (LoadSessionSnapshot(sessionFilePath, &snapshot) && !snapshot.history.empty()) {
                std::vector<std::pair<std::wstring, std::wstring>> additions;
                const bool collected = CollectSessionAutoPhraseAdditions(
                    snapshot,
                    primaryDictPath,
                    fallbackDictPath,
                    phraseSourceDictPath,
                    userDictPath,
                    extendDictPath,
                    {runtimeDictPath, helperDictPath},
                    additions);
                if (collected && !additions.empty() && !AppendAutoPhraseEntries(helperDictPath, additions)) {
                    std::cerr << "failed to append helper auto phrase dictionary\n";
                }
            }
        }

        const DWORD waitResult = WaitForSingleObject(wakeEvent, kSessionWatchPollIntervalMs);
        if (waitResult != WAIT_OBJECT_0 && waitResult != WAIT_TIMEOUT) {
            break;
        }
    }

    CloseHandle(wakeEvent);
    CloseHandle(helperMutex);
    return 0;
}

}  // namespace

int main(int argc, char* argv[]) {
    if (argc < 2) {
        std::cerr << "usage: yuninput_user_dict_builder <merge|extend|session-build> ...\n";
        return 1;
    }

    const std::string command = argv[1];
    if (command == "merge") {
        return RunMergeCommand(argc, argv);
    }
    if (command == "extend") {
        return RunExtendCommand(argc, argv);
    }
    if (command == "session-build") {
        return RunSessionBuildCommand(argc, argv);
    }
    if (command == "session-watch") {
        return RunSessionWatchCommand(argc, argv);
    }

    std::cerr << "unknown command: " << command << '\n';
    return 1;
}
