#include "Globals.h"
#include "TextService.h"

#include <Windows.h>
#include <msctf.h>
#include <olectl.h>

#include <fstream>
#include <string>

namespace {

bool SetRegStringValue(HKEY root, const std::wstring& subKey, const std::wstring& valueName, const std::wstring& data) {
    HKEY key = nullptr;
    if (RegCreateKeyExW(root, subKey.c_str(), 0, nullptr, REG_OPTION_NON_VOLATILE, KEY_WRITE, nullptr, &key, nullptr) != ERROR_SUCCESS) {
        return false;
    }

    const wchar_t* valueNamePtr = valueName.empty() ? nullptr : valueName.c_str();
    const DWORD bytes = static_cast<DWORD>((data.size() + 1) * sizeof(wchar_t));
    const LONG result = RegSetValueExW(
        key,
        valueNamePtr,
        0,
        REG_SZ,
        reinterpret_cast<const BYTE*>(data.c_str()),
        bytes);

    RegCloseKey(key);
    return result == ERROR_SUCCESS;
}

std::wstring GetRegistrationLogPath() {
    wchar_t localAppData[MAX_PATH] = {};
    const DWORD len = GetEnvironmentVariableW(L"LOCALAPPDATA", localAppData, MAX_PATH);
    if (len == 0 || len >= MAX_PATH) {
        return L"C:\\Windows\\Temp\\yuninput_register.log";
    }

    return std::wstring(localAppData, len) + L"\\yuninput\\register.log";
}

void AppendRegistrationLog(const std::wstring& message) {
    const std::wstring logPath = GetRegistrationLogPath();
    const std::size_t slash = logPath.find_last_of(L"\\/");
    if (slash != std::wstring::npos) {
        const std::wstring dirPath = logPath.substr(0, slash);
        CreateDirectoryW(dirPath.c_str(), nullptr);
    }

    std::ofstream stream(logPath, std::ios::app | std::ios::binary);
    if (!stream.is_open()) {
        return;
    }

    SYSTEMTIME now = {};
    GetLocalTime(&now);
    wchar_t prefix[64] = {};
    swprintf_s(
        prefix,
        L"[%04u-%02u-%02u %02u:%02u:%02u] ",
        now.wYear,
        now.wMonth,
        now.wDay,
        now.wHour,
        now.wMinute,
        now.wSecond);

    const std::wstring line = std::wstring(prefix) + message + L"\r\n";
    const unsigned char bom[] = {0xFF, 0xFE};
    if (stream.tellp() == 0) {
        stream.write(reinterpret_cast<const char*>(bom), sizeof(bom));
    }
    stream.write(reinterpret_cast<const char*>(line.data()), static_cast<std::streamsize>(line.size() * sizeof(wchar_t)));
}

void LogRegistrationHr(const wchar_t* step, HRESULT hr) {
    wchar_t buffer[160] = {};
    swprintf_s(buffer, L"%ls hr=0x%08X", step, static_cast<unsigned int>(hr));
    AppendRegistrationLog(buffer);
}

HRESULT RegisterComServer() {
    AppendRegistrationLog(L"RegisterComServer begin");
    std::wstring modulePath;
    if (!GetModulePath(modulePath)) {
        AppendRegistrationLog(L"GetModulePath failed");
        return E_FAIL;
    }

    const std::wstring clsid = GuidToString(CLSID_YunmaTextService);
    const std::wstring clsidKey = L"Software\\Classes\\CLSID\\" + clsid;
    const std::wstring inprocKey = clsidKey + L"\\InprocServer32";

    if (!SetRegStringValue(HKEY_LOCAL_MACHINE, clsidKey, L"", kTextServiceDescription)) {
        AppendRegistrationLog(L"Failed to write CLSID description");
        return E_FAIL;
    }

    if (!SetRegStringValue(HKEY_LOCAL_MACHINE, inprocKey, L"", modulePath)) {
        AppendRegistrationLog(L"Failed to write InprocServer32 path");
        return E_FAIL;
    }

    if (!SetRegStringValue(HKEY_LOCAL_MACHINE, inprocKey, L"ThreadingModel", L"Apartment")) {
        AppendRegistrationLog(L"Failed to write ThreadingModel");
        return E_FAIL;
    }

    AppendRegistrationLog(L"RegisterComServer success");
    return S_OK;
}

void UnregisterComServer() {
    const std::wstring clsid = GuidToString(CLSID_YunmaTextService);
    const std::wstring clsidKey = L"Software\\Classes\\CLSID\\" + clsid;
    RegDeleteTreeW(HKEY_LOCAL_MACHINE, clsidKey.c_str());
}

HRESULT RegisterProfiles() {
    AppendRegistrationLog(L"RegisterProfiles begin");
    HRESULT initHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    const bool shouldUninit = SUCCEEDED(initHr);
    if (initHr == RPC_E_CHANGED_MODE) {
        initHr = S_OK;
    }
    LogRegistrationHr(L"CoInitializeEx", initHr);

    ITfInputProcessorProfiles* profiles = nullptr;
    HRESULT hr = CoCreateInstance(
        CLSID_TF_InputProcessorProfiles,
        nullptr,
        CLSCTX_INPROC_SERVER,
        IID_ITfInputProcessorProfiles,
        reinterpret_cast<void**>(&profiles));

    if (FAILED(hr)) {
        LogRegistrationHr(L"CoCreateInstance(ITfInputProcessorProfiles)", hr);
        if (shouldUninit) {
            CoUninitialize();
        }
        return hr;
    }

    hr = profiles->Register(CLSID_YunmaTextService);
    if (hr == TF_E_ALREADY_EXISTS) {
        hr = S_OK;
    }
    LogRegistrationHr(L"ITfInputProcessorProfiles::Register", hr);

    if (SUCCEEDED(hr)) {
        std::wstring modulePath;
        if (GetModulePath(modulePath)) {
            hr = profiles->AddLanguageProfile(
                CLSID_YunmaTextService,
                kLangIdSimplifiedChinese,
                GUID_PROFILE_YUNMA,
                kLanguageProfileDescription,
                static_cast<ULONG>(wcslen(kLanguageProfileDescription)),
                modulePath.c_str(),
                static_cast<ULONG>(modulePath.size()),
                0);

            if (hr == TF_E_ALREADY_EXISTS) {
                hr = S_OK;
            }
            LogRegistrationHr(L"ITfInputProcessorProfiles::AddLanguageProfile", hr);

            if (SUCCEEDED(hr)) {
                hr = profiles->EnableLanguageProfile(
                    CLSID_YunmaTextService,
                    kLangIdSimplifiedChinese,
                    GUID_PROFILE_YUNMA,
                    TRUE);
                LogRegistrationHr(L"ITfInputProcessorProfiles::EnableLanguageProfile", hr);
            }
        } else {
            AppendRegistrationLog(L"GetModulePath failed before AddLanguageProfile");
            hr = E_FAIL;
        }
    }

    if (SUCCEEDED(hr)) {
        ITfCategoryMgr* categoryMgr = nullptr;
        HRESULT categoryHr = CoCreateInstance(
            CLSID_TF_CategoryMgr,
            nullptr,
            CLSCTX_INPROC_SERVER,
            IID_ITfCategoryMgr,
            reinterpret_cast<void**>(&categoryMgr));

        if (SUCCEEDED(categoryHr) && categoryMgr != nullptr) {
            const HRESULT categoryRegisterHr = categoryMgr->RegisterCategory(CLSID_YunmaTextService, GUID_TFCAT_TIP_KEYBOARD, CLSID_YunmaTextService);
            LogRegistrationHr(L"ITfCategoryMgr::RegisterCategory", categoryRegisterHr);
            categoryMgr->Release();
        } else {
            LogRegistrationHr(L"CoCreateInstance(ITfCategoryMgr)", categoryHr);
        }
    }

    profiles->Release();
    if (shouldUninit) {
        CoUninitialize();
    }

    LogRegistrationHr(L"RegisterProfiles result", hr);
    return hr;
}

void UnregisterProfiles() {
    HRESULT initHr = CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);
    const bool shouldUninit = SUCCEEDED(initHr);
    if (initHr == RPC_E_CHANGED_MODE) {
        initHr = S_OK;
    }

    ITfInputProcessorProfiles* profiles = nullptr;
    HRESULT hr = CoCreateInstance(
        CLSID_TF_InputProcessorProfiles,
        nullptr,
        CLSCTX_INPROC_SERVER,
        IID_ITfInputProcessorProfiles,
        reinterpret_cast<void**>(&profiles));

    if (SUCCEEDED(hr) && profiles != nullptr) {
        profiles->EnableLanguageProfile(
            CLSID_YunmaTextService,
            kLangIdSimplifiedChinese,
            GUID_PROFILE_YUNMA,
            FALSE);

        profiles->RemoveLanguageProfile(
            CLSID_YunmaTextService,
            kLangIdSimplifiedChinese,
            GUID_PROFILE_YUNMA);

        profiles->Unregister(CLSID_YunmaTextService);
        profiles->Release();
    }

    ITfCategoryMgr* categoryMgr = nullptr;
    hr = CoCreateInstance(
        CLSID_TF_CategoryMgr,
        nullptr,
        CLSCTX_INPROC_SERVER,
        IID_ITfCategoryMgr,
        reinterpret_cast<void**>(&categoryMgr));

    if (SUCCEEDED(hr) && categoryMgr != nullptr) {
        categoryMgr->UnregisterCategory(CLSID_YunmaTextService, GUID_TFCAT_TIP_KEYBOARD, CLSID_YunmaTextService);
        categoryMgr->Release();
    }

    if (shouldUninit) {
        CoUninitialize();
    }
}

}  // namespace

BOOL APIENTRY DllMain(HMODULE moduleHandle, DWORD reason, LPVOID) {
    if (reason == DLL_PROCESS_ATTACH) {
        g_moduleHandle = moduleHandle;
        DisableThreadLibraryCalls(moduleHandle);
    }

    return TRUE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv) {
    if (ppv == nullptr) {
        return E_INVALIDARG;
    }

    *ppv = nullptr;
    if (rclsid != CLSID_YunmaTextService) {
        return CLASS_E_CLASSNOTAVAILABLE;
    }

    return CreateTextServiceClassFactory(riid, ppv);
}

STDAPI DllCanUnloadNow() {
    return (g_dllRefCount == 0 && g_objectCount == 0) ? S_OK : S_FALSE;
}

STDAPI DllRegisterServer() {
    AppendRegistrationLog(L"DllRegisterServer begin");
    const HRESULT comHr = RegisterComServer();
    if (FAILED(comHr)) {
        LogRegistrationHr(L"RegisterComServer result", comHr);
        return SELFREG_E_CLASS;
    }

    const HRESULT profileHr = RegisterProfiles();
    if (FAILED(profileHr)) {
        LogRegistrationHr(L"RegisterProfiles result", profileHr);
        UnregisterComServer();
        return SELFREG_E_CLASS;
    }

    AppendRegistrationLog(L"DllRegisterServer success");
    return S_OK;
}

STDAPI DllUnregisterServer() {
    UnregisterProfiles();
    UnregisterComServer();
    return S_OK;
}
