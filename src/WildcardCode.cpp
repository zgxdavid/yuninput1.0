#include "WildcardCode.h"

namespace yuninput {

bool IsSupportedWildcardCodePattern(const std::wstring& code) {
    if (code.size() != 4) {
        return false;
    }

    bool hasWildcard = false;
    for (size_t index = 0; index < code.size(); ++index) {
        const wchar_t ch = code[index];
        if (ch == kWildcardCodeChar) {
            if (index != 1 && index != 3) {
                return false;
            }
            hasWildcard = true;
            continue;
        }

        if (ch < L'a' || ch > L'z') {
            return false;
        }
    }

    return hasWildcard;
}

void ExpandWildcardCodePattern(const std::wstring& pattern, std::vector<std::wstring>& outCodes) {
    outCodes.clear();
    if (!IsSupportedWildcardCodePattern(pattern)) {
        return;
    }

    std::wstring building = pattern;
    const bool wildcardAtSecond = building[1] == kWildcardCodeChar;
    const bool wildcardAtFourth = building[3] == kWildcardCodeChar;

    if (!wildcardAtSecond && !wildcardAtFourth) {
        outCodes.push_back(building);
        return;
    }

    if (wildcardAtSecond && wildcardAtFourth) {
        outCodes.reserve(26 * 26);
        for (wchar_t second = L'a'; second <= L'z'; ++second) {
            building[1] = second;
            for (wchar_t fourth = L'a'; fourth <= L'z'; ++fourth) {
                building[3] = fourth;
                outCodes.push_back(building);
            }
        }
        return;
    }

    outCodes.reserve(26);
    const size_t wildcardIndex = wildcardAtSecond ? 1 : 3;
    for (wchar_t ch = L'a'; ch <= L'z'; ++ch) {
        building[wildcardIndex] = ch;
        outCodes.push_back(building);
    }
}

}  // namespace yuninput
