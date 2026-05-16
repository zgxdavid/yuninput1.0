#include "CompositionEngine.h"

#include <Windows.h>

#include <clocale>
#include <iostream>
#include <string>

namespace {

LONGLONG QueryCounter() {
    LARGE_INTEGER value = {};
    QueryPerformanceCounter(&value);
    return value.QuadPart;
}

double ElapsedMs(LONGLONG startCounter, LONGLONG endCounter) {
    LARGE_INTEGER frequency = {};
    QueryPerformanceFrequency(&frequency);
    if (frequency.QuadPart <= 0) {
        return 0.0;
    }

    return (static_cast<double>(endCounter - startCounter) * 1000.0) / static_cast<double>(frequency.QuadPart);
}

std::wstring MakeCode(size_t index) {
    // Always produce a 4-letter code so it stays within normal zh input shape.
    wchar_t code[5] = {L'a', L'a', L'a', L'a', L'\0'};
    size_t v = index;
    for (int i = 3; i >= 0; --i) {
        code[i] = static_cast<wchar_t>(L'a' + (v % 26));
        v /= 26;
    }
    return std::wstring(code);
}

std::wstring MakeText(size_t index) {
    return L"perf_" + std::to_wstring(index);
}

}  // namespace

int wmain(int argc, wchar_t* argv[]) {
    if (argc < 3) {
        std::cerr << "usage: yuninput_index_insert_probe <dict-path> <insert-count>\n";
        return 1;
    }

    const std::wstring dictPath = argv[1];
    const size_t insertCount = static_cast<size_t>(std::wcstoul(argv[2], nullptr, 10));
    if (insertCount == 0) {
        std::cerr << "insert-count must be > 0\n";
        return 2;
    }

    SetConsoleOutputCP(CP_UTF8);
    SetConsoleCP(CP_UTF8);
    std::setlocale(LC_ALL, ".UTF-8");

    CompositionEngine engine;
    if (!engine.LoadDictionaryFromFile(dictPath)) {
        std::cerr << "failed to load dictionary\n";
        return 3;
    }

    size_t addedCount = 0;
    const LONGLONG startCounter = QueryCounter();
    for (size_t i = 0; i < insertCount; ++i) {
        const std::wstring code = MakeCode(i + 100000);
        const std::wstring text = MakeText(i);
        if (engine.AddUserEntry(code, text)) {
            ++addedCount;
        }
    }
    const LONGLONG endCounter = QueryCounter();

    const double elapsedMs = ElapsedMs(startCounter, endCounter);
    const double avgUs = insertCount > 0 ? (elapsedMs * 1000.0 / static_cast<double>(insertCount)) : 0.0;

    std::wcout << L"SUMMARY\t" << insertCount << L"\t" << addedCount << L"\t" << elapsedMs << L"\t" << avgUs << L"\n";
    return 0;
}
