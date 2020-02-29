#include <shlwapi.h>
#include <windows.h>

#include "myserver.h"

extern "C" {
#include "generate.h"
}

const TCHAR ProgIDStr[] = TEXT("Sample.MyServer2");
LONG LockCount{};
HINSTANCE MyServerDLLInstance{};
TCHAR MyServerCLSIDStr[256]{};
TCHAR LibraryIDStr[256]{};

// CMyServer
CMyServer::CMyServer()
    : mReferenceCount(1), mFile(nullptr), mTypeInfo(nullptr) {
  ITypeLib *pTypeLib{};
  HRESULT hr{};

  LockModule(TRUE);

  hr = LoadRegTypeLib(LIBID_MyServerLib, 1, 0, 0, &pTypeLib);

  if (SUCCEEDED(hr)) {
    pTypeLib->GetTypeInfoOfGuid(IID_IFileControl, &mTypeInfo);
    pTypeLib->Release();
  }
}

CMyServer::~CMyServer() {
  if (mTypeInfo != nullptr) {
    mTypeInfo->Release();
  }

  LockModule(FALSE);
}

STDMETHODIMP CMyServer::QueryInterface(REFIID riid, void **ppvObject) {
  *ppvObject = nullptr;
  if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IDispatch) ||
      IsEqualIID(riid, IID_IFileControl)) {
    *ppvObject = static_cast<IFileControl *>(this);
  } else {
    return E_NOINTERFACE;
  }

  AddRef();

  return S_OK;
}

STDMETHODIMP_(ULONG) CMyServer::AddRef() {
  return InterlockedIncrement(&mReferenceCount);
}

STDMETHODIMP_(ULONG) CMyServer::Release() {
  if (InterlockedDecrement(&mReferenceCount) == 0) {
    delete this;

    return 0;
  }

  return mReferenceCount;
}

STDMETHODIMP CMyServer::GetTypeInfoCount(UINT *pctinfo) {
  *pctinfo = mTypeInfo != nullptr ? 1 : 0;

  return S_OK;
}

STDMETHODIMP CMyServer::GetTypeInfo(UINT iTInfo, LCID lcid,
                                    ITypeInfo **ppTInfo) {
  if (mTypeInfo == nullptr) {
    return E_NOTIMPL;
  }

  mTypeInfo->AddRef();
  *ppTInfo = mTypeInfo;

  return S_OK;
}

STDMETHODIMP CMyServer::GetIDsOfNames(REFIID riid, LPOLESTR *rgszNames,
                                      UINT cNames, LCID lcid,
                                      DISPID *rgDispId) {
  if (mTypeInfo == nullptr) {
    return E_NOTIMPL;
  }

  return mTypeInfo->GetIDsOfNames(rgszNames, cNames, rgDispId);
}

STDMETHODIMP CMyServer::Invoke(DISPID dispIdMember, REFIID riid, LCID lcid,
                               WORD wFlags, DISPPARAMS *pDispParams,
                               VARIANT *pVarResult, EXCEPINFO *pExcepInfo,
                               UINT *puArgErr) {
  void *p = static_cast<IFileControl *>(this);

  if (mTypeInfo == nullptr) {
    return E_NOTIMPL;
  }

  return mTypeInfo->Invoke(p, dispIdMember, wFlags, pDispParams, pVarResult,
                           pExcepInfo, puArgErr);
}

STDMETHODIMP CMyServer::CreateFile(BSTR bstrFileName, FILEMODE filemode) {
  DWORD dwDesiredAccess;
  DWORD dwCreationDisposition;
  if (filemode == FM_READ) {
    dwDesiredAccess = GENERIC_READ;
    dwCreationDisposition = OPEN_EXISTING;
  } else if (filemode == FM_WRITE) {
    dwDesiredAccess = GENERIC_WRITE;
    dwCreationDisposition = CREATE_ALWAYS;
  } else {
    return E_FAIL;
  }

  mFile = ::CreateFileW(bstrFileName, dwDesiredAccess, 0, nullptr,
                        dwCreationDisposition, FILE_ATTRIBUTE_NORMAL, nullptr);

  return HRESULT_FROM_WIN32(GetLastError());
}

STDMETHODIMP CMyServer::ReadFile(DWORD dwLength, BSTR *lp) {
  DWORD dwReadByte{};
  LPSTR lpszA{};
  LPWSTR lpszW{};

  lpszA = (LPSTR)CoTaskMemAlloc(dwLength + 1);
  ::ReadFile(mFile, lpszA, dwLength, &dwReadByte, nullptr);
  lpszA[dwReadByte] = '\0';
  dwReadByte++;
  lpszW = (LPWSTR)CoTaskMemAlloc(dwReadByte * sizeof(WCHAR));
  MultiByteToWideChar(CP_ACP, 0, lpszA, -1, lpszW, dwReadByte);
  *lp = SysAllocString(lpszW);
  CoTaskMemFree(lpszA);
  CoTaskMemFree(lpszW);

  return S_OK;
}

STDMETHODIMP CMyServer::WriteFile(BSTR bstrData, DWORD dwLength) {
  DWORD dw, dwWriteByte{};
  LPSTR lpszA{};

  dw = SysStringLen(bstrData) + 1;
  lpszA = (LPSTR)CoTaskMemAlloc(dw);
  WideCharToMultiByte(CP_ACP, 0, bstrData, -1, lpszA, dw, nullptr, nullptr);
  ::WriteFile(mFile, lpszA, dwLength, &dwWriteByte, nullptr);

  CoTaskMemFree(lpszA);

  return HRESULT_FROM_WIN32(GetLastError());
}

STDMETHODIMP CMyServer::CloseFile() {
  CloseHandle(mFile);

  return HRESULT_FROM_WIN32(GetLastError());
}

STDMETHODIMP CMyServer::put_FilePos(DWORD dwPos) {
  SetFilePointer(mFile, dwPos, nullptr, FILE_CURRENT);

  return HRESULT_FROM_WIN32(GetLastError());
}

// CMyServerFactory
STDMETHODIMP CMyServerFactory::QueryInterface(REFIID riid, void **ppvObject) {
  *ppvObject = nullptr;

  if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IClassFactory)) {
    *ppvObject = static_cast<IClassFactory *>(this);
  } else {
    return E_NOINTERFACE;
  }

  AddRef();

  return S_OK;
}

STDMETHODIMP_(ULONG) CMyServerFactory::AddRef() {
  LockModule(TRUE);

  return 2;
}

STDMETHODIMP_(ULONG) CMyServerFactory::Release() {
  LockModule(FALSE);

  return 1;
}

STDMETHODIMP CMyServerFactory::CreateInstance(IUnknown *pUnkOuter, REFIID riid,
                                              void **ppvObject) {
  CMyServer *p{};
  HRESULT hr{};

  *ppvObject = nullptr;

  if (pUnkOuter != nullptr) {
    return CLASS_E_NOAGGREGATION;
  }

  p = new CMyServer();

  if (p == nullptr) {
    return E_OUTOFMEMORY;
  }

  hr = p->QueryInterface(riid, ppvObject);
  p->Release();

  return hr;
}

STDMETHODIMP CMyServerFactory::LockServer(BOOL fLock) {
  LockModule(fLock);

  return S_OK;
}

// DLL Export
STDAPI DllCanUnloadNow(void) { return LockCount == 0 ? S_OK : S_FALSE; }

STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID *ppv) {
  static CMyServerFactory serverFactory;
  HRESULT hr{};
  *ppv = nullptr;

  if (IsEqualCLSID(rclsid, CLSID_MyServer)) {
    hr = serverFactory.QueryInterface(riid, ppv);
  } else {
    hr = CLASS_E_CLASSNOTAVAILABLE;
  }

  return hr;
}

STDAPI DllRegisterServer(void) {
  TCHAR szModulePath[256]{};
  TCHAR szKey[256]{};
  WCHAR szTypeLibPath[256]{};
  HRESULT hr{};
  ITypeLib *pTypeLib{};

  wsprintf(szKey, TEXT("CLSID\\%s"), MyServerCLSIDStr);

  if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, nullptr,
                         TEXT("COM Server Sample2"))) {
    return E_FAIL;
  }

  GetModuleFileName(MyServerDLLInstance, szModulePath,
                    sizeof(szModulePath) / sizeof(TCHAR));
  wsprintf(szKey, TEXT("CLSID\\%s\\InprocServer32"), MyServerCLSIDStr);

  if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, nullptr, szModulePath)) {
    return E_FAIL;
  }

  wsprintf(szKey, TEXT("CLSID\\%s\\InprocServer32"), MyServerCLSIDStr);

  if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, TEXT("ThreadingModel"),
                         TEXT("Apartment"))) {
    return E_FAIL;
  }

  wsprintf(szKey, TEXT("CLSID\\%s\\ProgID"), MyServerCLSIDStr);

  if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, nullptr,
                         (LPTSTR)ProgIDStr)) {
    return E_FAIL;
  }

  wsprintf(szKey, TEXT("%s"), ProgIDStr);

  if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, nullptr,
                         TEXT("COM Server Sample2"))) {
    return E_FAIL;
  }

  wsprintf(szKey, TEXT("%s\\CLSID"), ProgIDStr);

  if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, nullptr,
                         (LPTSTR)MyServerCLSIDStr)) {
    return E_FAIL;
  }

  GetModuleFileNameW(MyServerDLLInstance, szTypeLibPath,
                     sizeof(szTypeLibPath) / sizeof(TCHAR));

  hr = LoadTypeLib(szTypeLibPath, &pTypeLib);

  if (SUCCEEDED(hr)) {
    hr = RegisterTypeLib(pTypeLib, szTypeLibPath, nullptr);

    if (SUCCEEDED(hr)) {
      wsprintf(szKey, TEXT("CLSID\\%s\\TypeLib"), MyServerCLSIDStr);

      if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, nullptr, LibraryIDStr)) {
        pTypeLib->Release();

        return E_FAIL;
      }
    }

    pTypeLib->Release();
  }

  return S_OK;
}

STDAPI DllUnregisterServer(void) {
  TCHAR szKey[256]{};

  wsprintf(szKey, TEXT("CLSID\\%s"), MyServerCLSIDStr);
  SHDeleteKey(HKEY_CLASSES_ROOT, szKey);

  wsprintf(szKey, TEXT("%s"), ProgIDStr);
  SHDeleteKey(HKEY_CLASSES_ROOT, szKey);

  UnRegisterTypeLib(LIBID_MyServerLib, 1, 0, LOCALE_NEUTRAL, SYS_WIN32);

  return S_OK;
}

BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved) {
  switch (dwReason) {
  case DLL_PROCESS_ATTACH:
    MyServerDLLInstance = hInstance;
    DisableThreadLibraryCalls(hInstance);
    GetGuidString(CLSID_MyServer, MyServerCLSIDStr);
    GetGuidString(LIBID_MyServerLib, LibraryIDStr);

    return TRUE;
  }

  return TRUE;
}

// Helper function
void LockModule(BOOL bLock) {
  if (bLock) {
    InterlockedIncrement(&LockCount);
  } else {
    InterlockedDecrement(&LockCount);
  }
}

BOOL CreateRegistryKey(HKEY hKeyRoot, LPTSTR lpszKey, LPTSTR lpszValue,
                       LPTSTR lpszData) {
  HKEY hKey{};
  LONG lResult{};
  DWORD dwSize{};

  lResult =
      RegCreateKeyEx(hKeyRoot, lpszKey, 0, nullptr, REG_OPTION_NON_VOLATILE,
                     KEY_WRITE, nullptr, &hKey, nullptr);

  if (lResult != ERROR_SUCCESS) {
    return FALSE;
  }
  if (lpszData != nullptr) {
    dwSize = (lstrlen(lpszData) + 1) * sizeof(TCHAR);
  } else {
    dwSize = 0;
  }

  RegSetValueEx(hKey, lpszValue, 0, REG_SZ, (LPBYTE)lpszData, dwSize);
  RegCloseKey(hKey);

  return TRUE;
}

void GetGuidString(REFGUID rguid, LPTSTR lpszGuid) {
  LPWSTR lpsz;
  StringFromCLSID(rguid, &lpsz);

#ifdef UNICODE
  lstrcpyW(lpszGuid, lpsz);
#else
  WideCharToMultiByte(CP_ACP, 0, lpsz, -1, lpszGuid, 256, nullptr, nullptr);
#endif

  CoTaskMemFree(lpsz);
}
