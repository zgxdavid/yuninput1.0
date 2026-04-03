#include "Globals.h"

#include <initguid.h>
#include <Objbase.h>

HMODULE g_moduleHandle = nullptr;
long g_dllRefCount = 0;
long g_objectCount = 0;

const wchar_t* kTextServiceDescription = L"yuninput Text Service";
const wchar_t* kLanguageProfileDescription = L"\u5300\u7801\u8f93\u5165\u6cd5";

// These GUIDs are project-local identifiers for the text service and language profile.
DEFINE_GUID(
    CLSID_YunmaTextService,
    0x6de9ab40,
    0x3ba8,
    0x4b77,
    0x8d,
    0x8f,
    0x23,
    0x39,
    0x66,
    0xe1,
    0xc1,
    0x02);

DEFINE_GUID(
    GUID_PROFILE_YUNMA,
    0x47de2fb1,
    0xf5e4,
    0x4cf8,
    0xab,
    0x2f,
    0x8f,
    0x7a,
    0x76,
    0x17,
    0x31,
    0xb2);

std::wstring GuidToString(REFGUID guid) {
    wchar_t buffer[64] = {};
    if (StringFromGUID2(guid, buffer, static_cast<int>(std::size(buffer))) <= 0) {
        return L"";
    }

    return std::wstring(buffer);
}

bool GetModulePath(std::wstring& outPath) {
    wchar_t path[MAX_PATH] = {};
    const DWORD len = GetModuleFileNameW(g_moduleHandle, path, MAX_PATH);
    if (len == 0 || len >= MAX_PATH) {
        return false;
    }

    outPath.assign(path, len);
    return true;
}

bool EnsureUserDataDirectory(std::wstring& outDir) {
    wchar_t dataRoot[MAX_PATH] = {};
    DWORD len = GetEnvironmentVariableW(L"APPDATA", dataRoot, MAX_PATH);
    if (len == 0 || len >= MAX_PATH) {
        len = GetEnvironmentVariableW(L"LOCALAPPDATA", dataRoot, MAX_PATH);
        if (len == 0 || len >= MAX_PATH) {
            return false;
        }
    }

    outDir = std::wstring(dataRoot, len) + L"\\yuninput";

    if (CreateDirectoryW(outDir.c_str(), nullptr) == 0) {
        const DWORD error = GetLastError();
        if (error != ERROR_ALREADY_EXISTS) {
            return false;
        }
    }

    return true;
}
