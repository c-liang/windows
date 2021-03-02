#pragma once
#include <combaseapi.h>

MIDL_INTERFACE("c04b329e-5823-4415-9c93-ba44688947b0")
IThemeManagerShared : public IUnknown {
 public:
  virtual /* [local] */ HRESULT STDMETHODCALLTYPE get_CurrentTheme(
      /* [annotation][iid_is][out] */
      _COM_Outptr_ void** ppvCurrentTheme) = 0;

  virtual /* [local] */ HRESULT STDMETHODCALLTYPE ApplyTheme(
      /* [in] */ BSTR themeFile) = 0;
};


MIDL_INTERFACE("d23cc733-5522-406d-8dfb-b3cf5ef52a71")
IThemeShared : public IUnknown {
 public:
  virtual /* [local] */ HRESULT STDMETHODCALLTYPE get_DisplayName(
      /* [annotation][iid_is][out] */
      _COM_Outptr_ WCHAR * *themeName) = 0;

  virtual /* [local] */ HRESULT STDMETHODCALLTYPE get_VisualStyle(
      /* [in] */ WCHAR * *visualStyle) = 0;
};