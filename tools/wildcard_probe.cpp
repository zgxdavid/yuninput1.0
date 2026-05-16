#include "CompositionEngine.h"
#include "WildcardCode.h"

#include <Windows.h>

#include <clocale>
#include <iostream>
#include <set>
#include <string>
#include <vector>

namespace {

std::string WideToUtf8(const std::wstring& input) {
    if (input.empty()) {
        return "";
    }

    const int required = WideCharToMultiByte(
        CP_UTF8,
        0,
        input.c_str(),
        static_cast<int>(input.size()),
        nullptr,
        0,
        nullptr,
        nullptr);
    if (required <= 0) {
        return "";
    }

    std::string output(static_cast<size_t>(required), '\0');
    WideCharToMultiByte(
        CP_UTF8,
        0,
        input.c_str(),
        static_cast<int>(input.size()),
        output.data(),
        required,
        nullptr,
        nullptr);
    return output;
}

}  // namespace

int wmain(int argc, wchar_t* argv[]) {
    if (argc < 3) {
        std::cerr << "usage: yuninput_wildcard_probe <dict-path> <pattern> [max-per-expanded-code]\n";
        return 1;
    }

    const std::wstring dictPath = argv[1];
    const std::wstring pattern = argv[2];
    size_t maxPerExpandedCode = 12;
    if (argc >= 4) {
        maxPerExpandedCode = static_cast<size_t>(std::wcstoul(argv[3], nullptr, 10));
        if (maxPerExpandedCode == 0) {
            maxPerExpandedCode = 12;
        }
    }

    SetConsoleOutputCP(CP_UTF8);
    SetConsoleCP(CP_UTF8);
    std::setlocale(LC_ALL, ".UTF-8");

    CompositionEngine engine;
    if (!engine.LoadDictionaryFromFile(dictPath)) {
        std::cerr << "failed to load dictionary\n";
        return 2;
    }

    if (!yuninput::IsSupportedWildcardCodePattern(pattern)) {
        std::cout << "SUMMARY\t" << WideToUtf8(pattern) << "\tINVALID\t0\t0\t0\n";
        return 3;
    }

    std::vector<std::wstring> expandedCodes;
    yuninput::ExpandWildcardCodePattern(pattern, expandedCodes);

    std::set<std::wstring> expandedCodesWithHits;
    size_t totalCandidateRows = 0;

    for (const std::wstring& expandedCode : expandedCodes) {
        const std::vector<CompositionEngine::Entry> matches =
            engine.QueryExactCandidateEntries(expandedCode, maxPerExpandedCode);
        if (!matches.empty()) {
            expandedCodesWithHits.insert(expandedCode);
        }

        std::cout << "EXPANDED\t" << WideToUtf8(expandedCode) << "\t" << matches.size() << "\n";
        for (const auto& entry : matches) {
            ++totalCandidateRows;
            std::cout << "CAND\t"
                      << WideToUtf8(expandedCode) << "\t"
                      << WideToUtf8(entry.code) << "\t"
                      << WideToUtf8(entry.text) << "\n";
        }
    }

    std::cout << "SUMMARY\t"
              << WideToUtf8(pattern) << "\tVALID\t"
              << expandedCodes.size() << "\t"
              << expandedCodesWithHits.size() << "\t"
              << totalCandidateRows << "\n";

    return 0;
}
