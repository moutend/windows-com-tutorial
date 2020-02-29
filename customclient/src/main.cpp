#include <windows.h>

#include "generate.h"

int main(Platform::Array<Platform::String ^> ^ args) {
  HRESULT hr{};
  IFileControl *pFileControl{};
  TCHAR szBuf[256]{};
  BSTR bstr{};

  CoInitialize(nullptr);

  hr = CoCreateInstance(CLSID_MyServer, nullptr, CLSCTX_INPROC_SERVER,
                        IID_PPV_ARGS(&pFileControl));

  if (FAILED(hr)) {
    MessageBox(nullptr, "Failed to create instance", nullptr, MB_ICONWARNING);
    CoUninitialize();
    return 0;
  }

  bstr = SysAllocString(L"C:\\hello.txt");
  hr = pFileControl->CreateFile(bstr, FM_READ);
  SysFreeString(bstr);

  if (FAILED(hr)) {
    MessageBox(nullptr, TEXT("Failed to open file."), nullptr, MB_ICONWARNING);
    pFileControl->Release();
    CoUninitialize();
    return 0;
  }

  pFileControl->ReadFile(3, &bstr);
  MessageBoxW(nullptr, bstr, L"OK", MB_OK);

  SysFreeString(bstr);
  pFileControl->CloseFile();
  pFileControl->Release();

  CoUninitialize();

  return 0;
}
