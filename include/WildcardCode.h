#pragma once

#include <string>
#include <vector>

namespace yuninput {

constexpr wchar_t kWildcardCodeChar = L'0';

bool IsSupportedWildcardCodePattern(const std::wstring& code);
void ExpandWildcardCodePattern(const std::wstring& pattern, std::vector<std::wstring>& outCodes);

}  // namespace yuninput
