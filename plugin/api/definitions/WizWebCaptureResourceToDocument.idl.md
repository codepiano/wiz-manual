```
// WizWebCaptureResourceToDocument.idl : IDL source for WizWebCaptureResourceToDocument
//

// This file will be processed by the MIDL tool to
// produce the type library (WizWebCaptureResourceToDocument.tlb) and marshalling code.

import "oaidl.idl";
import "ocidl.idl";

[
    object,
    uuid(E28ABCEE-E1E2-4819-AE67-27A38608D63D),
    dual,
    nonextensible,
    helpstring("IWizProtectedIEHelper Interface"),
    pointer_default(unique)
]
interface IWizProtectedIEHelper : IDispatch{
    [id(1), helpstring("method SaveDocument")] HRESULT SaveDocument([in] IDispatch* pCaller, [in] IDispatch* pDocDisp, [in] BSTR bstrResourceFileName);
};
[
    uuid(326D8F1F-AC46-4C2F-9676-185AF800C695),
    version(1.0),
    helpstring("WizWebCaptureResourceToDocument 1.0 Type Library")
]
library WizWebCaptureResourceToDocumentLib
{
    importlib("stdole2.tlb");
    [
        uuid(A43FD1D4-FF5F-4532-9CD1-690366C9E339),
        helpstring("WizProtectedIEHelper Class")
    ]
    coclass WizProtectedIEHelper
    {
        [default] interface IWizProtectedIEHelper;
    };
};
```
