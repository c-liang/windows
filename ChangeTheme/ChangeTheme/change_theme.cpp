// ChangeTheme.cpp : This file contains the 'main' function. Program execution
// begins and ends there.
//

#include <Windows.h>

#include <ShlObj.h>
#include <combaseapi.h>
#include <wrl/client.h>

#include <filesystem>
#include <iostream>

#include "theme_traits.h"

HRESULT get_current_theme_path(std::wstring& theme_path) {
  HRESULT hr = E_NOTIMPL;
  do {
    HKEY key_result;
    auto status =
        RegOpenKeyExW(HKEY_CURRENT_USER,
                      L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes",
                      0, KEY_READ, &key_result);
    if (status != ERROR_SUCCESS) {
      hr = HRESULT_FROM_WIN32(GetLastError());
      break;
    }

    DWORD byte_size = 2 * MAX_PATH * sizeof(wchar_t);
    wchar_t local_mem[2 * MAX_PATH]{0};
    status = RegQueryValueExW(key_result, L"CurrentTheme", nullptr, nullptr,
                              (LPBYTE)&local_mem[0], &byte_size);
    if (status != ERROR_SUCCESS) {
      hr = HRESULT_FROM_WIN32(GetLastError());
      break;
    }
    theme_path = local_mem;
  } while (false);
  return hr;
}

HRESULT apply_theme(const wchar_t* theme_file) {
  std::error_code err;
  if (!std::filesystem::exists(theme_file, err)) {
    return HRESULT_FROM_WIN32(ERROR_FILE_NOT_FOUND);
  }
  Microsoft::WRL::ComPtr<IClassFactory> factory;

  HRESULT hr =
      CoGetClassObject(__uuidof(IThemeManagerShared), CLSCTX_INPROC_SERVER,
                       nullptr, IID_IClassFactory, &factory);
  if (FAILED(hr)) {
    return hr;
  }
  Microsoft::WRL::ComPtr<IThemeManagerShared> theme_manager_shared;
  hr = factory->CreateInstance(nullptr, IID_IUnknown, &theme_manager_shared);
  if (FAILED(hr)) {
    return hr;
  }
  BSTR new_theme_file = ::SysAllocString(theme_file);
  if (!new_theme_file) {
    return HRESULT_FROM_WIN32(ERROR_OUTOFMEMORY);
  }
  hr = theme_manager_shared->ApplyTheme(new_theme_file);
  ::SysFreeString(new_theme_file);
  return hr;
}

HRESULT get_display_name(std::wstring& theme_name) {
  Microsoft::WRL::ComPtr<IClassFactory> factory;

  HRESULT hr =
      CoGetClassObject(__uuidof(IThemeManagerShared), CLSCTX_INPROC_SERVER,
                       nullptr, IID_IClassFactory, &factory);
  if (FAILED(hr)) {
    return hr;
  }
  Microsoft::WRL::ComPtr<IThemeManagerShared> theme_manager_shared;
  hr = factory->CreateInstance(nullptr, IID_IUnknown, &theme_manager_shared);
  if (FAILED(hr)) {
    return hr;
  }
  Microsoft::WRL::ComPtr<IThemeShared> theme_shared;
  theme_manager_shared->get_CurrentTheme(&theme_shared);
  if (FAILED(hr)) {
    return hr;
  }
  WCHAR* display_name;
  hr = theme_shared->get_DisplayName(&display_name);
  if (SUCCEEDED(hr)) {
    theme_name = display_name;
  }
  return hr;
}
HRESULT get_visual_style(std::wstring& visual_style) {
  Microsoft::WRL::ComPtr<IClassFactory> factory;

  HRESULT hr =
      CoGetClassObject(__uuidof(IThemeManagerShared), CLSCTX_INPROC_SERVER,
                       nullptr, IID_IClassFactory, &factory);
  if (FAILED(hr)) {
    return hr;
  }
  Microsoft::WRL::ComPtr<IThemeManagerShared> theme_manager_shared;
  hr = factory->CreateInstance(nullptr, IID_IUnknown, &theme_manager_shared);
  if (FAILED(hr)) {
    return hr;
  }
  Microsoft::WRL::ComPtr<IThemeShared> theme_shared;
  theme_manager_shared->get_CurrentTheme(&theme_shared);
  if (FAILED(hr)) {
    return hr;
  }
  WCHAR* style_name;
  hr = theme_shared->get_VisualStyle(&style_name);
  if (SUCCEEDED(hr)) {
    visual_style = style_name;
  }
  return hr;
}
int main() {
  HRESULT hr = CoInitialize(nullptr);
  std::wstring current_theme_path;
  get_current_theme_path(current_theme_path);

  std::wstring display_name;
  hr = get_display_name(display_name);
  std::wcout << display_name;
  std::cout << hr << std::endl;
  std::wstring visual_style;
  hr = get_visual_style(visual_style);
  std::wcout << visual_style;
  std::cout << hr << std::endl;
  apply_theme(L"C:\\Windows\\Resources\\Themes\\aero.theme");
  std::cout << hr << std::endl;

  // apply_theme(current_theme_path.c_str());
  CoUninitialize();
  system("pause");
}

// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started:
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add
//   Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project
//   and select the .sln file
