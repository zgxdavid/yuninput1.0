#pragma once

#include <Windows.h>
#include <guiddef.h>

#include <string>

extern HMODULE g_moduleHandle;
extern long g_dllRefCount;
extern long g_objectCount;

extern const CLSID CLSID_YunmaTextService;
extern const GUID GUID_PROFILE_YUNMA;

extern const wchar_t* kTextServiceDescription;
extern const wchar_t* kLanguageProfileDescription;

constexpr LANGID kLangIdSimplifiedChinese = MAKELANGID(LANG_CHINESE, SUBLANG_CHINESE_SIMPLIFIED);

std::wstring GuidToString(REFGUID guid);
bool GetModulePath(std::wstring& outPath);
bool EnsureUserDataDirectory(std::wstring& outDir);
