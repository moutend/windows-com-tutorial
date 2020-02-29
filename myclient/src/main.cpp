#include <windows.h>

const CLSID CLSID_MyServer = {0x112143a6,
                              0x62c1,
                              0x4478,
                              {0x9e, 0x8f, 0x87, 0x26, 0x99, 0x25, 0x5e, 0x2e}};

int main(Platform::Array<Platform::String ^> ^ args) {
  HRESULT hr{};
  CLSID clsid{};
  OLECHAR *pszProgId{};
  IPersistFile *pPersistFile{};
  ISequentialStream *pSequentialStream{};
  ULONG uRead{};
  TCHAR szBuf[256]{};
  char szData[256]{};

  CoInitialize(nullptr);

  hr = ProgIDFromCLSID(CLSID_MyServer, &pszProgId);

  if (FAILED(hr)) {
    MessageBox(nullptr, TEXT("Failed to convert from CLSID to ProgID."),
               nullptr, MB_ICONWARNING);

    CoUninitialize();

    return 0;
  }

  hr = CLSIDFromProgID(pszProgId, &clsid);

  if (FAILED(hr)) {
    MessageBox(nullptr, TEXT("Failed to convert from ProgID to CLSID."),
               nullptr, MB_ICONWARNING);

    CoUninitialize();

    return 0;
  }

  hr = CoCreateInstance(CLSID_MyServer, NULL, CLSCTX_INPROC_SERVER,
                        IID_PPV_ARGS(&pPersistFile));

  if (FAILED(hr)) {
    MessageBox(nullptr, TEXT("Failed to create instance."), nullptr,
               MB_ICONWARNING);

    CoUninitialize();

    return 0;
  }

  /*
    hr = pPersistFile->Load(TEXT("C:\\sample.txt"), STGM_READ);

    if (FAILED(hr)) {
      MessageBox(nullptr, TEXT("Failed to open file."), nullptr,
                 MB_ICONWARNING);

      pPersistFile->Release();

      CoUninitialize();

      return 0;
    }

    hr = pPersistFile->QueryInterface(IID_PPV_ARGS(&pSequentialStream));

    if (FAILED(hr)) {
      MessageBox(nullptr, TEXT("Failed to query ISequentialStream."), nullptr,
    MB_ICONWARNING);

      pPersistFile->Release();

      CoUninitialize();

      return 0;
    }

    hr = pSequentialStream->Read(szData, 5, &uRead);
    szData[uRead] = '\0';

    MessageBoxA(nullptr, TEXT("Success"), "OK", MB_OK);

    pSequentialStream->Release();
*/

  if (pPersistFile->Release() != 0) {
    MessageBox(nullptr, TEXT("Failed to release object."), nullptr,
               MB_ICONWARNING);

    CoUninitialize();

    return 0;
  }

  MessageBox(nullptr, TEXT("Done."), TEXT("OK"), MB_OK);

  CoUninitialize();

  return 0;
}
