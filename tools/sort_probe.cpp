#include "CompositionEngine.h"

#include <Windows.h>

#include <cstdlib>
#include <iostream>
#include <string>
#include <vector>
#include <clocale>

namespace {

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

}  // namespace

int wmain(int argc, wchar_t* argv[]) {
    if (argc < 3) {
        std::cerr << "usage: yuninput_sort_probe <dict-path> <code> [max-candidates] [--record <code> <text> [boost]]...\n";
        std::cerr << "   or: yuninput_sort_probe <dict-path> --phrase <text>\n";
        return 1;
    }

    const std::wstring dictPath = argv[1];
    const std::wstring mode = argv[2];
    const bool phraseMode = mode == L"--phrase";
    const std::wstring code = phraseMode ? L"" : std::wstring(argv[2]);
    size_t maxCandidates = 10;
    int argIndex = 3;
    if (!phraseMode && argc >= 4 && std::wstring(argv[3]).rfind(L"--", 0) != 0) {
        maxCandidates = static_cast<size_t>(std::wcstoul(argv[3], nullptr, 10));
        argIndex = 4;
    }

    SetConsoleOutputCP(CP_UTF8);
    SetConsoleCP(CP_UTF8);
    std::setlocale(LC_ALL, ".UTF-8");

    CompositionEngine engine;
    if (!engine.LoadDictionaryFromFile(dictPath)) {
        std::cerr << "failed to load dictionary\n";
        return 2;
    }

    if (phraseMode) {
        if (argc < 4) {
            std::cerr << "--phrase requires <text>\n";
            return 3;
        }

        std::vector<std::wstring> phraseCodes;
        if (!engine.TryBuildPhraseCodes(std::wstring(argv[3]), phraseCodes) || phraseCodes.empty()) {
            std::cerr << "no phrase code generated\n";
            return 4;
        }

        for (const std::wstring& phraseCode : phraseCodes) {
            std::cout << WideToUtf8(phraseCode) << '\n';
        }
        return 0;
    }

    while (argIndex < argc) {
        const std::wstring option = argv[argIndex];
        if (option != L"--record") {
            std::cerr << "unknown option: " << WideToUtf8(option) << '\n';
            return 5;
        }
        if (argIndex + 2 >= argc) {
            std::cerr << "--record requires <code> <text> [boost]\n";
            return 6;
        }

        const std::wstring recordCode = argv[argIndex + 1];
        const std::wstring recordText = argv[argIndex + 2];
        std::uint64_t boost = 1;
        argIndex += 3;
        if (argIndex < argc && std::wstring(argv[argIndex]).rfind(L"--", 0) != 0) {
            boost = static_cast<std::uint64_t>(_wcstoui64(argv[argIndex], nullptr, 10));
            ++argIndex;
        }

        engine.RecordCommit(recordCode, recordText, boost);
    }

    const std::vector<CompositionEngine::Entry> candidates = engine.QueryCandidateEntries(code, maxCandidates);
    for (const auto& entry : candidates) {
        std::cout << WideToUtf8(entry.code) << '\t' << WideToUtf8(entry.text) << '\n';
    }

    return 0;
}