```
// TestCom.idl : IDL source for TestCom.exe
//

// This file will be processed by the MIDL tool to
// produce the type library (TestCom.tlb) and marshalling code.

import "oaidl.idl";
import "ocidl.idl";

[
    object,
    uuid(43C4CF61-5D88-4beb-B593-793A2251A491),
    dual,
    nonextensible,
    helpstring("IWizExplorerApp Interface"),
    pointer_default(unique)
]
interface IWizExplorerApp : IDispatch{
    [propget, id(1), helpstring("property Database")] HRESULT Database([out, retval] IDispatch** pVal);
    [propget, id(2), helpstring("property Window")] HRESULT Window([out, retval] IDispatch** pVal);
    [propget, id(3), helpstring("property AppPath")] HRESULT AppPath([out, retval] BSTR* pVal);
    [propget, id(4), helpstring("property DataStore")] HRESULT DataStore([out, retval] BSTR* pVal);
    [propget, id(5), helpstring("property SettingsFileName")] HRESULT SettingsFileName([out, retval] BSTR* pVal);
    [propget, id(6), helpstring("property LogFileName")] HRESULT LogFileName([out, retval] BSTR* pVal);
    [id(7), helpstring("method LoadPluginString")] HRESULT LoadPluginString([in] BSTR bstrPluginAppGUID, [in] BSTR bstrStringName, [out, retval] BSTR* pVal);
    [id(8), helpstring("method LoadPluginString2")] HRESULT LoadPluginString2([in] IDispatch* pHtmlDocumentDisp, [in] BSTR bstrStringName, [out, retval] BSTR* pVal);
    [id(9), helpstring("method TranslateString")] HRESULT TranslateString([in] BSTR bstrString, [out,retval] BSTR* pbstrRet);
    [id(10), helpstring("method PluginLocalizeHtmlDialog")] HRESULT PluginLocalizeHtmlDialog([in] IDispatch* pHtmlDocumentDisp);
    [id(11), helpstring("method GetPluginAppGUID")] HRESULT GetPluginAppGUID([in] IDispatch* pHtmlDocumentDisp, [out,retval] BSTR* pVal);
    [id(12), helpstring("method GetPluginAppPath")] HRESULT GetPluginAppPath([in] IDispatch* pHtmlDocumentDisp, [out,retval] BSTR* pVal);
    [id(13), helpstring("method CreateWizObject")] HRESULT CreateWizObject([in] BSTR bstrObjectName, [out,retval] IDispatch** ppVal);
    [id(14), helpstring("method CreateActiveXObject")] HRESULT CreateActiveXObject([in] BSTR bstrObjectName, [out,retval] IDispatch** ppVal);
    [propget, id(15), helpstring("property CurPluginAppGUID")] HRESULT CurPluginAppGUID([out, retval] BSTR* pVal);
    [propget, id(16), helpstring("property CurPluginGUID")] HRESULT CurPluginGUID([out, retval] BSTR* pVal);
    [propget, id(17), helpstring("property CurPluginAppPath")] HRESULT CurPluginAppPath([out, retval] BSTR* pVal);
    [id(18), helpstring("method LoadString")] HRESULT LoadString([in] BSTR bstrStringName, [out, retval] BSTR* pVal);
    [id(19), helpstring("method RunScriptCode")] HRESULT RunScriptCode([in] BSTR bstrScriptCode, [in] BSTR bstrScriptLanguage);
    [id(20), helpstring("method RunScriptFile")] HRESULT RunScriptFile([in] BSTR bstrScriptFileName, [in] BSTR bstrScriptLanguage);
    [id(21), helpstring("method LoadStringFromFile")] HRESULT LoadStringFromFile([in] BSTR bstrFileName, [in] BSTR bstrStringName, [out, retval] BSTR* pVal);
    [id(22), helpstring("method LocalizeHtmlDocument")] HRESULT LocalizeHtmlDocument([in] BSTR bstrLanguageFileName, [in] IDispatch* pHtmlDocumentDisp);
    [id(23), helpstring("method GetHtmlDocumentPath")] HRESULT GetHtmlDocumentPath([in] IDispatch* pHtmlDocumentDisp, [out,retval] BSTR* pVal);
    [id(24), helpstring("method GetPluginPathByScriptFileName")] HRESULT GetPluginPathByScriptFileName([in] BSTR bstrScriptFileName, [out,retval] BSTR* pbstrPluginPath);
    [id(25), helpstring("method AddGlobalScript")] HRESULT AddGlobalScript([in] BSTR bstrScriptFileName);
    [id(26), helpstring("method ExecutePlugin")] HRESULT ExecutePlugin([in] BSTR bstrPluginGUID, [in] BSTR bstrURLParams);
    [id(27), helpstring("method ExecuteCommand")] HRESULT ExecuteCommand([in] BSTR bstrCommandName, [in] VARIANT* pvInParam1, [in] VARIANT* pvInParam2, [out, retval] VARIANT* pvretResult);
    [id(28), helpstring("method InsertPluginMenu")] HRESULT InsertPluginMenu([in] LONGLONG nMenuHandle, [in] LONG nInsertPos, [in] BSTR bstrMenuType, [in] LONG nPluginBeginCommand, [out, retval] BSTR* pbstrCommandPluginGUID);
    [propget, id(29), helpstring("property SyncProgress")] HRESULT SyncProgress([out, retval] IDispatch** pVal);
    [id(30), helpstring("method GetGroupDatabase")] HRESULT GetGroupDatabase([in] BSTR bstrKbGUID, [out, retval] IDispatch** pVal);
};

[
    object,
    uuid(B15879A3-8647-4d6f-84D8-3A24709EC962),
    dual,
    nonextensible,
    helpstring("IWizExplorerWindow Interface"),
    pointer_default(unique)
]
interface IWizExplorerWindow : IDispatch{
    [propget, id(1), helpstring("property CategoryCtrl")] HRESULT CategoryCtrl([out, retval] IDispatch** pVal);
    [propget, id(2), helpstring("property DocumentsCtrl")] HRESULT DocumentsCtrl([out, retval] IDispatch** pVal);
    [propget, id(3), helpstring("property HWND")] HRESULT HWND([out, retval] OLE_HANDLE* pVal);
    [propget, id(4), helpstring("property CurrentDocument")] HRESULT CurrentDocument([out, retval] IDispatch** pVal);
    [propget, id(5), helpstring("property CurrentDocumentHtmlDocument")] HRESULT CurrentDocumentHtmlDocument([out, retval] IDispatch** pVal);
    [propget, id(6), helpstring("property CurrentDocumentAttachmentCtrl")] HRESULT CurrentDocumentAttachmentCtrl([out, retval] IDispatch** pVal);
    [id(7), helpstring("method ShowHtmlDialog")] HRESULT ShowHtmlDialog([in] BSTR bstrTitle, [in] BSTR bstrURL, [in] LONG nWidth, [in] LONG nHeight, [in] BSTR bstrExtOptions, [in] VARIANT vParam, [out, retval] VARIANT* pvRet);
    [id(8), helpstring("method CloseHtmlDialog")] HRESULT CloseHtmlDialog([in] IDispatch* pHtmlDocumentDisp, [in] VARIANT vRet);
    [id(9), helpstring("method GetHtmlDialogParam")] HRESULT GetHtmlDialogParam([in] IDispatch* pHtmlDocumentDisp, [out, retval] VARIANT* pvParam);
    [id(12), helpstring("method ViewDocument")] HRESULT ViewDocument([in] IDispatch* pDocumentDisp, [in] VARIANT_BOOL vbOpenInNewTab);
    [id(13), helpstring("method ViewHtml")] HRESULT ViewHtml([in] BSTR bstrURL, [in] VARIANT_BOOL vbOpenInNewTab);
    [id(14), helpstring("method GetHtmlDialogHWND")] HRESULT GetHtmlDialogHWND([in] IDispatch* pHtmlDocumentDisp, [out,retval] OLE_HANDLE* phHWND);
    [id(15), helpstring("method ShowMessage")] HRESULT ShowMessage([in] BSTR bstrText, [in] BSTR bstrTitle, [in] LONG nType, [out,retval] LONG* pnRet);
    [id(16), helpstring("method ShowSelectorWindow")] HRESULT ShowSelectorWindow([in] BSTR bstrURL, [in] LONG left, [in] LONG top, [in] LONG width, [in] LONG height, [in] BSTR bstrExtOptions);
    [id(17), helpstring("method CloseSelectorWindow")] HRESULT CloseSelectorWindow([in] IDispatch* pdispSelectorHtmlDocument);
    [id(18), helpstring("method AddToolButton")] HRESULT AddToolButton([in] BSTR bstrType, [in] BSTR bstrButtonID, [in] BSTR bstrButtonText, [in] BSTR bstrIconFileName, [in] BSTR bstrClickEventFunction);
    [id(19), helpstring("method RemoveToolButton")] HRESULT RemoveToolButton([in] BSTR bstrType, [in] BSTR bstrButtonID);
    [id(20), helpstring("method SetToolButtonState")] HRESULT SetToolButtonState([in] BSTR bstrType, [in] BSTR bstrButtonID, [in] BSTR bstrState);
    [id(21), helpstring("method GetToolButtonRect")] HRESULT GetToolButtonRect([in] BSTR bstrType, [in] BSTR bstrButtonID, [out,retval] BSTR* pbstrRect);
    [id(22), helpstring("method AddTimer")] HRESULT AddTimer([in] BSTR bstrCallback, [in] LONG nElapse);
    [id(23), helpstring("method RemoveTimer")] HRESULT RemoveTimer([in] BSTR bstrCallback);
    [id(24), helpstring("method ExecCommand")] HRESULT ExecCommand([in] BSTR bstrCommandName, [in] BSTR bstrParams);
};

[
    uuid(CA4A7BC1-2FB5-41ad-A2DD-12A1EDED6F28),
    version(1.0),
    helpstring("WizExplorer 1.0 Type Library")
]
library WizExplorerLib
{
    importlib("stdole32.tlb");
    importlib("stdole2.tlb");
    [
        uuid(BB2A9D7C-459A-486e-90EC-A8C80727FDD9),
        helpstring("WizExplorerApp Class")
    ]
    coclass WizExplorerApp
    {
        [default] interface IWizExplorerApp;
    };
    [
        uuid(C428B079-7D18-4ca3-BE64-E3F4A25300C5),
        helpstring("WizExplorerWindow Class")
    ]
    coclass WizExplorerWindow
    {
        [default] interface IWizExplorerWindow;
    };

};
```
