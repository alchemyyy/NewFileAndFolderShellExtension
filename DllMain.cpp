#include "NewFileAndFolderShellExtension.h"
#include <strsafe.h>

HINSTANCE g_hInst = NULL;
long g_cDllRef = 0;

BOOL APIENTRY DllMain(HMODULE hModule, DWORD dwReason, LPVOID lpReserved)
{
    switch (dwReason)
    {
    case DLL_PROCESS_ATTACH:
        g_hInst = hModule;
        DisableThreadLibraryCalls(hModule);
        break;
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
        break;
    }
    return TRUE;
}

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, void **ppv)
{
    HRESULT hr = CLASS_E_CLASSNOTAVAILABLE;

    if (IsEqualCLSID(CLSID_NewFileAndFolderShellExtension, rclsid))
    {
        hr = E_OUTOFMEMORY;

        ClassFactory *pClassFactory = new (std::nothrow) ClassFactory();
        if (pClassFactory)
        {
            hr = pClassFactory->QueryInterface(riid, ppv);
            pClassFactory->Release();
        }
    }

    return hr;
}

STDAPI DllCanUnloadNow(void)
{
    return g_cDllRef > 0 ? S_FALSE : S_OK;
}

STDAPI DllRegisterServer(void)
{
    HRESULT hr;

    WCHAR szModule[MAX_PATH];
    if (GetModuleFileName(g_hInst, szModule, ARRAYSIZE(szModule)) == 0)
    {
        hr = HRESULT_FROM_WIN32(GetLastError());
        return hr;
    }

    // Convert CLSID to string
    WCHAR szCLSID[MAX_PATH];
    StringFromGUID2(CLSID_NewFileAndFolderShellExtension, szCLSID, ARRAYSIZE(szCLSID));

    // Register CLSID
    WCHAR szSubkey[MAX_PATH];
    hr = StringCchPrintf(szSubkey, ARRAYSIZE(szSubkey), L"CLSID\\%s", szCLSID);
    if (SUCCEEDED(hr))
    {
        hr = HRESULT_FROM_WIN32(SHSetValue(HKEY_CLASSES_ROOT, szSubkey, NULL, 
            REG_SZ, L"NewFileAndFolderShellExtension", 
            sizeof(L"NewFileAndFolderShellExtension")));
    }

    if (SUCCEEDED(hr))
    {
        WCHAR szInprocServer32[MAX_PATH];
        hr = StringCchPrintf(szInprocServer32, ARRAYSIZE(szInprocServer32), 
            L"CLSID\\%s\\InprocServer32", szCLSID);
        if (SUCCEEDED(hr))
        {
            hr = HRESULT_FROM_WIN32(SHSetValue(HKEY_CLASSES_ROOT, szInprocServer32, 
                NULL, REG_SZ, szModule, (lstrlen(szModule) + 1) * sizeof(WCHAR)));

            if (SUCCEEDED(hr))
            {
                hr = HRESULT_FROM_WIN32(SHSetValue(HKEY_CLASSES_ROOT, szInprocServer32, 
                    L"ThreadingModel", REG_SZ, L"Apartment", 
                    sizeof(L"Apartment")));
            }
        }
    }

    // Register as context menu handler for directories and directory background
    if (SUCCEEDED(hr))
    {
        hr = HRESULT_FROM_WIN32(SHSetValue(HKEY_CLASSES_ROOT, 
            L"Directory\\Background\\shellex\\ContextMenuHandlers\\NewFileAndFolderShellExtension",
            NULL, REG_SZ, szCLSID, (lstrlen(szCLSID) + 1) * sizeof(WCHAR)));
    }

    if (SUCCEEDED(hr))
    {
        hr = HRESULT_FROM_WIN32(SHSetValue(HKEY_CLASSES_ROOT, 
            L"Directory\\shellex\\ContextMenuHandlers\\NewFileAndFolderShellExtension",
            NULL, REG_SZ, szCLSID, (lstrlen(szCLSID) + 1) * sizeof(WCHAR)));
    }

    return hr;
}

STDAPI DllUnregisterServer(void)
{
    HRESULT hr = S_OK;

    WCHAR szCLSID[MAX_PATH];
    StringFromGUID2(CLSID_NewFileAndFolderShellExtension, szCLSID, ARRAYSIZE(szCLSID));

    WCHAR szSubkey[MAX_PATH];
    hr = StringCchPrintf(szSubkey, ARRAYSIZE(szSubkey), L"CLSID\\%s", szCLSID);
    if (SUCCEEDED(hr))
    {
        SHDeleteKey(HKEY_CLASSES_ROOT, szSubkey);
    }

    SHDeleteKey(HKEY_CLASSES_ROOT, 
        L"Directory\\Background\\shellex\\ContextMenuHandlers\\NewFileAndFolderShellExtension");
    SHDeleteKey(HKEY_CLASSES_ROOT, 
        L"Directory\\shellex\\ContextMenuHandlers\\NewFileAndFolderShellExtension");

    return hr;
}
