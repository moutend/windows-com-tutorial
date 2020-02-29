#pragma once

#include <objidl.h>
#include <windows.h>

extern void LockModule(BOOL bLock);
extern BOOL CreateRegistryKey(HKEY hKeyRoot, LPTSTR lpszKey, LPTSTR lpszValue,
                              LPTSTR lpszData);

class CMyServer : public IPersistFile, public ISequentialStream {
public:
  // IUnknown methods.
  STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
  STDMETHODIMP_(ULONG) AddRef();
  STDMETHODIMP_(ULONG) Release();

  // IPersist method
  STDMETHODIMP GetClassID(CLSID *pClassID);

  // IPersistFile methods
  STDMETHODIMP IsDirty();
  STDMETHODIMP Load(LPCOLESTR pszFileName, DWORD dwMode);
  STDMETHODIMP Save(LPCOLESTR pszFileName, BOOL fRemember);
  STDMETHODIMP SaveCompleted(LPCOLESTR pszFileName);
  STDMETHODIMP GetCurFile(LPOLESTR *ppszFileName);

  // ISequentialStream methods
  STDMETHODIMP Read(void *pv, ULONG cb, ULONG *pcbRead);
  STDMETHODIMP Write(const void *pv, ULONG cb, ULONG *pcbWritten);

  CMyServer();
  ~CMyServer();

private:
  LONG mReferenceCount;
  HANDLE mFile;
};

class CMyServerFactory : public IClassFactory {
public:
  // IUnknown methods
  STDMETHODIMP QueryInterface(REFIID riid, void **ppvObject);
  STDMETHODIMP_(ULONG) AddRef();
  STDMETHODIMP_(ULONG) Release();

  // IClassFactory methods
  STDMETHODIMP CreateInstance(IUnknown *pUnkOuter, REFIID riid,
                              void **ppvObject);
  STDMETHODIMP LockServer(BOOL fLock);
};
