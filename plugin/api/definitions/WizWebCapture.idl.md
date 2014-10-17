```
// WizWebCapture.idl : IDL source for WizWebCapture
//

// This file will be processed by the MIDL tool to
// produce the type library (WizWebCapture.tlb) and marshalling code.

import "oaidl.idl";
import "ocidl.idl";

[
    object,
    uuid(502644EE-7301-457B-A77D-8BC3329AEA21),
    dual,
    nonextensible,
    helpstring("IWizIECapture Interface"),
    pointer_default(unique)
]
interface IWizIECapture : IDispatch{
    [id(1), helpstring("method Save")] HRESULT Save([in] IDispatch* pHtmlDocument2Disp, [in] VARIANT_BOOL vbSaveSel, [out,retval] IDispatch** ppWizDocumentDisp);
    [id(2), helpstring("method BackgroundSave")] HRESULT BackgroundSave([in] IDispatch* pHtmlDocument2Disp, [in] VARIANT_BOOL vbSaveSel);
    [propget, id(3), helpstring("property DatabasePath")] HRESULT DatabasePath([out, retval] BSTR* pVal);
    [propput, id(3), helpstring("property DatabasePath")] HRESULT DatabasePath([in] BSTR newVal);
    [propget, id(4), helpstring("property SaveButtonOnly")] HRESULT SaveButtonOnly([out, retval] VARIANT_BOOL* pVal);
    [propput, id(4), helpstring("property SaveButtonOnly")] HRESULT SaveButtonOnly([in] VARIANT_BOOL newVal);
    [id(5), helpstring("method Save2")] HRESULT Save2([in] IDispatch* pHtmlDocument2Disp);
};
[
    uuid(3293C7A3-79FF-44FE-BBC5-AD47E5E40A09),
    version(1.0),
    helpstring("WizWebCapture 1.0 Type Library")
]
library WizWebCaptureLib
{
    importlib("stdole2.tlb");
    [
        uuid(ED20B72B-C691-4331-A8D0-FBBB231A5DD4),
        helpstring("WizIECapture Class")
    ]
    coclass WizIECapture
    {
        [default] interface IWizIECapture;
    };
};
```
