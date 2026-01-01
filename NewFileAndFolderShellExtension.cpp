#include "NewFileAndFolderShellExtension.h"
#include <strsafe.h>
#include <new>

#pragma comment(lib, "shlwapi.lib")

extern long g_cDllRef;

#define IDM_NEW_FOLDER    0
#define IDM_NEW_TEXTFILE  1

NewFileAndFolderShellExtension::NewFileAndFolderShellExtension() 
    : m_cRef(1), m_hFolderIcon(NULL), m_hTextFileBitmap(NULL), m_pSite(NULL), m_pidlFolder(NULL)
{
    InterlockedIncrement(&g_cDllRef);
}

NewFileAndFolderShellExtension::~NewFileAndFolderShellExtension()
{
    if (m_hFolderIcon)
        DestroyIcon(m_hFolderIcon);
    if (m_hTextFileBitmap)
        DeleteObject(m_hTextFileBitmap);
    if (m_pSite)
        m_pSite->Release();
    if (m_pidlFolder)
        ILFree(m_pidlFolder);
    
    InterlockedDecrement(&g_cDllRef);
}

#pragma region IUnknown

IFACEMETHODIMP NewFileAndFolderShellExtension::QueryInterface(REFIID riid, void **ppv)
{
    static const QITAB qit[] = 
    {
        QITABENT(NewFileAndFolderShellExtension, IContextMenu),
        QITABENT(NewFileAndFolderShellExtension, IShellExtInit),
        QITABENT(NewFileAndFolderShellExtension, IObjectWithSite),
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP_(ULONG) NewFileAndFolderShellExtension::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

IFACEMETHODIMP_(ULONG) NewFileAndFolderShellExtension::Release()
{
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (0 == cRef)
    {
        delete this;
    }
    return cRef;
}

#pragma endregion

#pragma region IObjectWithSite

IFACEMETHODIMP NewFileAndFolderShellExtension::SetSite(IUnknown *pUnkSite)
{
    if (m_pSite)
    {
        m_pSite->Release();
        m_pSite = NULL;
    }
    
    if (pUnkSite)
    {
        m_pSite = pUnkSite;
        m_pSite->AddRef();
    }
    
    return S_OK;
}

IFACEMETHODIMP NewFileAndFolderShellExtension::GetSite(REFIID riid, void **ppvSite)
{
    if (!ppvSite)
        return E_POINTER;
    
    *ppvSite = NULL;
    
    if (!m_pSite)
        return E_FAIL;
    
    return m_pSite->QueryInterface(riid, ppvSite);
}

#pragma endregion

#pragma region IShellExtInit

IFACEMETHODIMP NewFileAndFolderShellExtension::Initialize(
    LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hKeyProgID)
{
    if (pidlFolder)
    {
        // Store a copy of the folder PIDL
        if (m_pidlFolder)
        {
            ILFree(m_pidlFolder);
            m_pidlFolder = NULL;
        }
        m_pidlFolder = ILClone(pidlFolder);
        
        WCHAR szFolder[MAX_PATH];
        if (SHGetPathFromIDList(pidlFolder, szFolder))
        {
            m_szFolder = szFolder;
            return S_OK;
        }
    }
    
    return E_FAIL;
}

#pragma endregion

#pragma region IContextMenu

IFACEMETHODIMP NewFileAndFolderShellExtension::QueryContextMenu(
    HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    if (CMF_DEFAULTONLY & uFlags)
    {
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(0));
    }

    // Get the folder icon from shell32.dll (same as Windows uses)
    SHFILEINFO sfi = {0};
    SHGetFileInfo(L"C:\\", 0, &sfi, sizeof(sfi), SHGFI_ICON | SHGFI_SMALLICON);
    m_hFolderIcon = sfi.hIcon;

    // Get the text file icon and convert to 32-bit bitmap with alpha
    SHGetFileInfo(L".txt", FILE_ATTRIBUTE_NORMAL, &sfi, sizeof(sfi), 
                  SHGFI_ICON | SHGFI_SMALLICON | SHGFI_USEFILEATTRIBUTES);
    if (sfi.hIcon)
    {
        int cx = GetSystemMetrics(SM_CXSMICON);
        int cy = GetSystemMetrics(SM_CYSMICON);
        
        // Create 32-bit DIB section for alpha support
        BITMAPINFO bmi = {0};
        bmi.bmiHeader.biSize = sizeof(BITMAPINFOHEADER);
        bmi.bmiHeader.biWidth = cx;
        bmi.bmiHeader.biHeight = -cy; // top-down
        bmi.bmiHeader.biPlanes = 1;
        bmi.bmiHeader.biBitCount = 32;
        bmi.bmiHeader.biCompression = BI_RGB;
        
        HDC hdcScreen = GetDC(NULL);
        HDC hdcMem = CreateCompatibleDC(hdcScreen);
        void *pvBits;
        m_hTextFileBitmap = CreateDIBSection(hdcScreen, &bmi, DIB_RGB_COLORS, &pvBits, NULL, 0);
        HBITMAP hOldBmp = (HBITMAP)SelectObject(hdcMem, m_hTextFileBitmap);
        
        // Clear to transparent
        memset(pvBits, 0, cx * cy * 4);
        
        DrawIconEx(hdcMem, 0, 0, sfi.hIcon, cx, cy, 0, NULL, DI_NORMAL);
        SelectObject(hdcMem, hOldBmp);
        DeleteDC(hdcMem);
        ReleaseDC(NULL, hdcScreen);
        DestroyIcon(sfi.hIcon);
    }

    // Add "New Folder" menu item
    MENUITEMINFO mii = {0};
    mii.cbSize = sizeof(mii);
    mii.fMask = MIIM_BITMAP | MIIM_STRING | MIIM_FTYPE | MIIM_ID | MIIM_STATE;
    mii.wID = idCmdFirst + IDM_NEW_FOLDER;
    mii.fType = MFT_STRING;
    mii.dwTypeData = (LPWSTR)L"New Folder";
    mii.fState = MFS_ENABLED;
    mii.hbmpItem = HBMMENU_CALLBACK;
    
    if (!InsertMenuItem(hMenu, indexMenu, TRUE, &mii))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    // Add "New Text File" menu item
    mii.wID = idCmdFirst + IDM_NEW_TEXTFILE;
    mii.dwTypeData = (LPWSTR)L"New Text Document";
    mii.hbmpItem = m_hTextFileBitmap;
    
    if (!InsertMenuItem(hMenu, indexMenu + 1, TRUE, &mii))
    {
        return HRESULT_FROM_WIN32(GetLastError());
    }

    return MAKE_HRESULT(SEVERITY_SUCCESS, 0, USHORT(2));
}

IFACEMETHODIMP NewFileAndFolderShellExtension::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
    BOOL fUnicode = FALSE;

    if (pici->cbSize == sizeof(CMINVOKECOMMANDINFOEX))
    {
        if (pici->fMask & CMIC_MASK_UNICODE)
        {
            fUnicode = TRUE;
        }
    }

    if (!fUnicode && HIWORD(pici->lpVerb))
    {
        if (StrCmpIA(pici->lpVerb, "newfolder") == 0)
        {
            CreateNewFolder();
            return S_OK;
        }
        else if (StrCmpIA(pici->lpVerb, "newtextfile") == 0)
        {
            CreateNewTextFile();
            return S_OK;
        }
        return E_FAIL;
    }

    if (fUnicode && HIWORD(((CMINVOKECOMMANDINFOEX*)pici)->lpVerbW))
    {
        if (StrCmpIW(((CMINVOKECOMMANDINFOEX*)pici)->lpVerbW, L"newfolder") == 0)
        {
            CreateNewFolder();
            return S_OK;
        }
        else if (StrCmpIW(((CMINVOKECOMMANDINFOEX*)pici)->lpVerbW, L"newtextfile") == 0)
        {
            CreateNewTextFile();
            return S_OK;
        }
        return E_FAIL;
    }

    if (LOWORD(pici->lpVerb) == IDM_NEW_FOLDER)
    {
        CreateNewFolder();
        return S_OK;
    }
    else if (LOWORD(pici->lpVerb) == IDM_NEW_TEXTFILE)
    {
        CreateNewTextFile();
        return S_OK;
    }

    return E_FAIL;
}

IFACEMETHODIMP NewFileAndFolderShellExtension::GetCommandString(
    UINT_PTR idCmd, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax)
{
    HRESULT hr = E_INVALIDARG;

    if (idCmd == IDM_NEW_FOLDER)
    {
        switch (uFlags)
        {
        case GCS_VERBA:
            hr = StringCchCopyA(pszName, cchMax, "newfolder");
            break;
        case GCS_VERBW:
            hr = StringCchCopyW((LPWSTR)pszName, cchMax, L"newfolder");
            break;
        case GCS_HELPTEXTA:
            hr = StringCchCopyA(pszName, cchMax, "Create a new folder");
            break;
        case GCS_HELPTEXTW:
            hr = StringCchCopyW((LPWSTR)pszName, cchMax, L"Create a new folder");
            break;
        }
    }
    else if (idCmd == IDM_NEW_TEXTFILE)
    {
        switch (uFlags)
        {
        case GCS_VERBA:
            hr = StringCchCopyA(pszName, cchMax, "newtextfile");
            break;
        case GCS_VERBW:
            hr = StringCchCopyW((LPWSTR)pszName, cchMax, L"newtextfile");
            break;
        case GCS_HELPTEXTA:
            hr = StringCchCopyA(pszName, cchMax, "Create a new text file");
            break;
        case GCS_HELPTEXTW:
            hr = StringCchCopyW((LPWSTR)pszName, cchMax, L"Create a new text file");
            break;
        }
    }

    return hr;
}

#pragma endregion

#pragma region Helper Methods

void NewFileAndFolderShellExtension::CreateNewFolder()
{
    if (m_szFolder.empty())
        return;

    std::wstring baseName = L"New folder";
    std::wstring folderName = baseName;
    std::wstring folderPath = m_szFolder + L"\\" + folderName;
    int counter = 1;

    // Find unique name
    while (PathFileExists(folderPath.c_str()))
    {
        WCHAR szCounter[32];
        StringCchPrintf(szCounter, ARRAYSIZE(szCounter), L" (%d)", counter++);
        folderName = baseName + szCounter;
        folderPath = m_szFolder + L"\\" + folderName;
    }

    // Create the folder
    if (CreateDirectory(folderPath.c_str(), NULL))
    {
        // Notify the shell that a new folder was created
        SHChangeNotify(SHCNE_MKDIR, SHCNF_PATH | SHCNF_FLUSH, folderPath.c_str(), NULL);
        
        SelectAndRename(folderName);
    }
}

void NewFileAndFolderShellExtension::CreateNewTextFile()
{
    if (m_szFolder.empty())
        return;

    std::wstring baseName = L"New Text Document";
    std::wstring fileName = baseName + L".txt";
    std::wstring filePath = m_szFolder + L"\\" + fileName;
    int counter = 1;

    // Find unique name
    while (PathFileExists(filePath.c_str()))
    {
        WCHAR szCounter[32];
        StringCchPrintf(szCounter, ARRAYSIZE(szCounter), L" (%d)", counter++);
        fileName = baseName + szCounter + L".txt";
        filePath = m_szFolder + L"\\" + fileName;
    }

    // Create the file
    HANDLE hFile = CreateFile(filePath.c_str(), GENERIC_WRITE, 0, NULL, 
                              CREATE_NEW, FILE_ATTRIBUTE_NORMAL, NULL);
    if (hFile != INVALID_HANDLE_VALUE)
    {
        CloseHandle(hFile);
        
        // Notify the shell that a new file was created
        SHChangeNotify(SHCNE_CREATE, SHCNF_PATH | SHCNF_FLUSH, filePath.c_str(), NULL);
        
        SelectAndRename(fileName);
    }
}

void NewFileAndFolderShellExtension::SelectAndRename(const std::wstring& itemName)
{
    if (!m_pSite)
        return;

    // Get IServiceProvider from the site
    IServiceProvider *pServiceProvider = NULL;
    HRESULT hr = m_pSite->QueryInterface(IID_IServiceProvider, (void**)&pServiceProvider);
    if (FAILED(hr) || !pServiceProvider)
        return;

    // Get IShellBrowser from service provider
    IShellBrowser *pShellBrowser = NULL;
    hr = pServiceProvider->QueryService(SID_STopLevelBrowser, IID_IShellBrowser, (void**)&pShellBrowser);
    if (FAILED(hr) || !pShellBrowser)
    {
        pServiceProvider->Release();
        return;
    }

    // Get IShellView from shell browser
    IShellView *pShellView = NULL;
    hr = pShellBrowser->QueryActiveShellView(&pShellView);
    if (FAILED(hr) || !pShellView)
    {
        pShellBrowser->Release();
        pServiceProvider->Release();
        return;
    }

    // Get IFolderView from shell view
    IFolderView *pFolderView = NULL;
    hr = pShellView->QueryInterface(IID_IFolderView, (void**)&pFolderView);
    if (FAILED(hr) || !pFolderView)
    {
        pShellView->Release();
        pShellBrowser->Release();
        pServiceProvider->Release();
        return;
    }

    // Get IShellFolder from folder view
    IShellFolder *pShellFolder = NULL;
    hr = pFolderView->GetFolder(IID_IShellFolder, (void**)&pShellFolder);
    if (SUCCEEDED(hr) && pShellFolder)
    {
        // Parse the item name to get its PIDL (relative to the folder)
        LPITEMIDLIST pidlItem = NULL;
        hr = pShellFolder->ParseDisplayName(NULL, NULL, (LPWSTR)itemName.c_str(), NULL, &pidlItem, NULL);
        if (SUCCEEDED(hr) && pidlItem)
        {
            // Select the item and enter rename mode
            pShellView->SelectItem(pidlItem, SVSI_SELECT | SVSI_DESELECTOTHERS | SVSI_ENSUREVISIBLE | SVSI_FOCUSED | SVSI_EDIT);
            
            CoTaskMemFree(pidlItem);
        }
        pShellFolder->Release();
    }

    pFolderView->Release();
    pShellView->Release();
    pShellBrowser->Release();
    pServiceProvider->Release();
}

#pragma endregion

#pragma region ClassFactory

ClassFactory::ClassFactory() : m_cRef(1)
{
    InterlockedIncrement(&g_cDllRef);
}

ClassFactory::~ClassFactory()
{
    InterlockedDecrement(&g_cDllRef);
}

IFACEMETHODIMP ClassFactory::QueryInterface(REFIID riid, void **ppv)
{
    static const QITAB qit[] = 
    {
        QITABENT(ClassFactory, IClassFactory),
        { 0 },
    };
    return QISearch(this, qit, riid, ppv);
}

IFACEMETHODIMP_(ULONG) ClassFactory::AddRef()
{
    return InterlockedIncrement(&m_cRef);
}

IFACEMETHODIMP_(ULONG) ClassFactory::Release()
{
    ULONG cRef = InterlockedDecrement(&m_cRef);
    if (0 == cRef)
    {
        delete this;
    }
    return cRef;
}

IFACEMETHODIMP ClassFactory::CreateInstance(
    IUnknown *pUnkOuter, REFIID riid, void **ppv)
{
    HRESULT hr;
    NewFileAndFolderShellExtension *pExt;

    *ppv = NULL;

    if (pUnkOuter != NULL)
    {
        return CLASS_E_NOAGGREGATION;
    }

    pExt = new (std::nothrow) NewFileAndFolderShellExtension();
    if (!pExt)
    {
        return E_OUTOFMEMORY;
    }

    hr = pExt->QueryInterface(riid, ppv);
    pExt->Release();

    return hr;
}

IFACEMETHODIMP ClassFactory::LockServer(BOOL fLock)
{
    if (fLock)
    {
        InterlockedIncrement(&g_cDllRef);
    }
    else
    {
        InterlockedDecrement(&g_cDllRef);
    }
    return S_OK;
}

#pragma endregion
