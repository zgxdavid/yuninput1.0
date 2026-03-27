#include "CompositionEngine.h"

#include <Windows.h>

#include <fcntl.h>
#include <iostream>
#include <io.h>
#include <string>
#include <vector>

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

    _setmode(_fileno(stdout), _O_U16TEXT);

    CompositionEngine engine;
    if (!engine.LoadDictionaryFromFile(dictPath)) {
        std::cerr << "failed to load dictionary\n";
        return 2;
    }

    const std::vector<CompositionEngine::Entry> candidates = engine.QueryCandidateEntries(code, maxCandidates);
    for (const auto& entry : candidates) {
        std::wcout << entry.code << L'\t' << entry.text << L'\n';
    }

    return 0;
}