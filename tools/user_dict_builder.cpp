#include "CompositionEngine.h"

#include <Windows.h>

#include <algorithm>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <set>
#include <sstream>
#include <string>
#include <unordered_set>
#include <vector>

namespace {

struct DictEntry {
    std::wstring code;
    std::wstring text;
    size_t sourcePriority = 0;
    size_t sourceOrder = 0;
};

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

}  // namespace

int main(int argc, char* argv[]) {
    if (argc < 2) {
        std::cerr << "usage: yuninput_user_dict_builder <merge|extend> ...\n";
        return 1;
    }

    const std::string command = argv[1];
    if (command == "merge") {
        return RunMergeCommand(argc, argv);
    }
    if (command == "extend") {
        return RunExtendCommand(argc, argv);
    }

    std::cerr << "unknown command: " << command << '\n';
    return 1;
}
