#pragma once

#include <windows.h>
#include <shlobj.h>
#include <shlwapi.h>
#include <shobjidl.h>
#include <exdisp.h>
#include <string>

// {8F8B6C1D-9E3A-4F2B-A1C5-D7E8F9A0B1C2}
static const GUID CLSID_NewFileAndFolderShellExtension = 
{ 0x8f8b6c1d, 0x9e3a, 0x4f2b, { 0xa1, 0xc5, 0xd7, 0xe8, 0xf9, 0xa0, 0xb1, 0xc2 } };

class NewFileAndFolderShellExtension : public IShellExtInit, public IContextMenu, public IObjectWithSite
{
public:
    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // IShellExtInit
    IFACEMETHODIMP Initialize(LPCITEMIDLIST pidlFolder, LPDATAOBJECT pDataObj, HKEY hKeyProgID);

    // IContextMenu
    IFACEMETHODIMP QueryContextMenu(HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags);
    IFACEMETHODIMP InvokeCommand(LPCMINVOKECOMMANDINFO pici);
    IFACEMETHODIMP GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax);

    // IObjectWithSite
    IFACEMETHODIMP SetSite(IUnknown *pUnkSite);
    IFACEMETHODIMP GetSite(REFIID riid, void **ppvSite);

    NewFileAndFolderShellExtension();

protected:
    ~NewFileAndFolderShellExtension();

private:
    long m_cRef;
    std::wstring m_szFolder;
    LPITEMIDLIST m_pidlFolder;
    IUnknown *m_pSite;
    
    HICON m_hFolderIcon;
    HBITMAP m_hTextFileBitmap;

    void CreateNewFolder();
    void CreateNewTextFile();
    void SelectAndRename(const std::wstring& itemName);
};

class ClassFactory : public IClassFactory
{
public:
    // IUnknown
    IFACEMETHODIMP QueryInterface(REFIID riid, void **ppv);
    IFACEMETHODIMP_(ULONG) AddRef();
    IFACEMETHODIMP_(ULONG) Release();

    // IClassFactory
    IFACEMETHODIMP CreateInstance(IUnknown *pUnkOuter, REFIID riid, void **ppv);
    IFACEMETHODIMP LockServer(BOOL fLock);

    ClassFactory();

protected:
    ~ClassFactory();

private:
    long m_cRef;
};
