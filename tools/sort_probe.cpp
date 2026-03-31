#include "CompositionEngine.h"

#include <Windows.h>

#include <iostream>
#include <string>
#include <vector>
#include <clocale>

namespace {

std::wstring Utf8ToWide(const std::string& input) {
    if (input.empty()) {
        return L"";
    }

    const int required = MultiByteToWideChar(CP_UTF8, 0, input.c_str(), static_cast<int>(input.size()), nullptr, 0);
    if (required <= 0) {
        return L"";
    }

    std::wstring output(static_cast<size_t>(required), L'\0');
    MultiByteToWideChar(CP_UTF8, 0, input.c_str(), static_cast<int>(input.size()), output.data(), required);
    return output;
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

int main(int argc, char* argv[]) {
    if (argc < 3) {
        std::cerr << "usage: yuninput_sort_probe <dict-path> <code> [max-candidates]\n";
        return 1;
    }

    const std::wstring dictPath = Utf8ToWide(argv[1]);
    const std::wstring code = Utf8ToWide(argv[2]);
    size_t maxCandidates = 10;
    if (argc >= 4) {
        maxCandidates = static_cast<size_t>(std::stoul(argv[3]));
    }

    SetConsoleOutputCP(CP_UTF8);
    SetConsoleCP(CP_UTF8);
    std::setlocale(LC_ALL, ".UTF-8");

    CompositionEngine engine;
    if (!engine.LoadDictionaryFromFile(dictPath)) {
        std::cerr << "failed to load dictionary\n";
        return 2;
    }

    const std::vector<CompositionEngine::Entry> candidates = engine.QueryCandidateEntries(code, maxCandidates);
    for (const auto& entry : candidates) {
        std::cout << WideToUtf8(entry.code) << '\t' << WideToUtf8(entry.text) << '\n';
    }

    return 0;
}