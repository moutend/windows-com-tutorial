#include <shlwapi.h>
#include <windows.h>

#include "myserver.h"

const CLSID CLSID_MyServer = {0x112143a6,
                              0x62c1,
                              0x4478,
                              {0x9e, 0x8f, 0x87, 0x26, 0x99, 0x25, 0x5e, 0x2e}};
const TCHAR MyServerCLSIDStr[] = TEXT("{112143A6-62C1-4478-9E8F-872699255E2E}");
const TCHAR MyServerProgIDStr[] = TEXT("Sample.MyServer");

LONG LockCount{};
HINSTANCE MyServerDLLInstance{};

// Helper functions
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

// CMyServer methods
CMyServer::CMyServer() : mReferenceCount(1), mFile(INVALID_HANDLE_VALUE) {
  LockModule(TRUE);
}

CMyServer::~CMyServer() {
  if (mFile != INVALID_HANDLE_VALUE) {
    CloseHandle(mFile);
  }

  LockModule(FALSE);
}

STDMETHODIMP CMyServer::QueryInterface(REFIID riid, void **ppvObject) {
  *ppvObject = nullptr;

  if (IsEqualIID(riid, IID_IUnknown) || IsEqualIID(riid, IID_IPersist) ||
      IsEqualIID(riid, IID_IPersistFile)) {
    *ppvObject = static_cast<IPersistFile *>(this);
  } else if (IsEqualIID(riid, IID_ISequentialStream)) {
    *ppvObject = static_cast<ISequentialStream *>(this);
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

STDMETHODIMP CMyServer::GetClassID(CLSID *pClassID) {
  // TODO: implement me!
  return E_NOTIMPL;
}

STDMETHODIMP CMyServer::IsDirty() {
  // TODO: implement me!
  return E_NOTIMPL;
}

STDMETHODIMP CMyServer::Load(LPCOLESTR pszFileName, DWORD dwMode) {
  DWORD dwError{};
  DWORD dwDesiredAccess{};
  DWORD dwCreationDisposition{};

  if (dwMode & STGM_READWRITE) {
    dwDesiredAccess = GENERIC_READ | GENERIC_WRITE;
  } else if (dwMode & STGM_WRITE) {
    dwDesiredAccess = GENERIC_WRITE;
  } else {
    dwDesiredAccess = GENERIC_READ;
  }
  if (dwMode & STGM_CREATE) {
    dwCreationDisposition = CREATE_ALWAYS;
  } else {
    dwCreationDisposition = OPEN_EXISTING;
    mFile = CreateFile(reinterpret_cast<LPCSTR>(pszFileName), dwDesiredAccess, 0, nullptr,
                       dwCreationDisposition, FILE_ATTRIBUTE_NORMAL, nullptr);
    dwError = GetLastError();
  }

  return HRESULT_FROM_WIN32(dwError);
}

STDMETHODIMP CMyServer::Save(LPCOLESTR pszFileName, BOOL fRemember) {
  // TODO: implement me!
  return E_NOTIMPL;
}

STDMETHODIMP CMyServer::SaveCompleted(LPCOLESTR pszFileName) {
  // TODO: implement me!
  return E_NOTIMPL;
}

STDMETHODIMP CMyServer::GetCurFile(LPOLESTR *ppszFileName) {
  // TODO: implement me!
  return E_NOTIMPL;
}

STDMETHODIMP CMyServer::Read(void *pv, ULONG cb, ULONG *pcbRead) {
  DWORD dwError{};
  ReadFile(mFile, pv, cb, pcbRead, nullptr);
  dwError = GetLastError();

  return HRESULT_FROM_WIN32(dwError);
}

STDMETHODIMP CMyServer::Write(const void *pv, ULONG cb, ULONG *pcbWritten) {
  // TODO: implement me!
  return E_NOTIMPL;
}

// CMyServerFactory methods
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

// Exported functions
STDAPI DllCanUnloadNow() { return LockCount == 0 ? S_OK : S_FALSE; }

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
  TCHAR szModulePath[MAX_PATH]{};
  TCHAR szKey[256]{};
  wsprintf(szKey, TEXT("CLSID\\%s"), MyServerCLSIDStr);

  if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, nullptr,
                         TEXT("COM Server Sample"))) {
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
                         (LPTSTR)MyServerProgIDStr)) {
    return E_FAIL;
  }

  wsprintf(szKey, TEXT("%s"), MyServerProgIDStr);

  if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, nullptr,
                         TEXT("COM Server Sample"))) {
    return E_FAIL;
  }

  wsprintf(szKey, TEXT("%s\\CLSID"), MyServerProgIDStr);

  if (!CreateRegistryKey(HKEY_CLASSES_ROOT, szKey, nullptr,
                         (LPTSTR)MyServerCLSIDStr)) {
    return E_FAIL;
  }

  return S_OK;
}

STDAPI DllUnregisterServer(void) {
  TCHAR szKey[256]{};

  wsprintf(szKey, TEXT("CLSID\\%s"), MyServerCLSIDStr);
  SHDeleteKey(HKEY_CLASSES_ROOT, szKey);

  wsprintf(szKey, TEXT("%s"), MyServerProgIDStr);
  SHDeleteKey(HKEY_CLASSES_ROOT, szKey);
  return S_OK;
}

BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved) {
  switch (dwReason) {
  case DLL_PROCESS_ATTACH:
    MyServerDLLInstance = hInstance;
    DisableThreadLibraryCalls(hInstance);
    return TRUE;
  }

  return TRUE;
}
