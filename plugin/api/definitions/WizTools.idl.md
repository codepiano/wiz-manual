```
import "oaidl.idl";
import "ocidl.idl";


[
    object,
    uuid(EF44AB0B-8E66-4EA7-B0CD-433336F47A2E),
    dual,
    nonextensible,
    helpstring("IWizHtmlEditorApp Interface"),
    pointer_default(unique)
]
interface IWizHtmlEditorApp : IDispatch{
    [propget, id(1), helpstring("property EditorDocument")] HRESULT EditorDocument([out, retval] IDispatch** pVal);
    [id(2), helpstring("method LocalizeHtmlDocument")] HRESULT LocalizeHtmlDocument([in] BSTR bstrLanguageFileName, [in] IDispatch* pHtmlDocumentDisp);
    [id(3), helpstring("method GetHtmlDocumentPath")] HRESULT GetHtmlDocumentPath([in] IDispatch* pHtmlDocumentDisp, [out,retval] BSTR* pVal);
    [id(4), helpstring("method CreateCommonTools")] HRESULT CreateCommonTools([out,retval] IDispatch** pVal);
    [id(5), helpstring("method ShowPluginDlg")] HRESULT ShowPluginDlg([in] BSTR bstrURL, LONG nWidth, LONG nHeight, BSTR bstrExtOptions);
    [id(6), helpstring("method ClosePluginDlg")] HRESULT ClosePluginDlg([in] IDispatch* pdispHtmlDocument, [in] LONG nRet);
    [id(7), helpstring("method ShowSelectorWindow")] HRESULT ShowSelectorWindow([in] BSTR bstrURL, [in] LONG left, [in] LONG top, [in] LONG width, [in] LONG height, [in] BSTR bstrExtOptions);
    [id(8), helpstring("method CloseSelectorWindow")] HRESULT CloseSelectorWindow([in] IDispatch* pdispSelectorHtmlDocument);
    [id(9), helpstring("method ShowMessage")] HRESULT ShowMessage([in] BSTR bstrText, [in] LONG nMsgBoxFlags, [out,retval] LONG* pnRet);
    [id(10), helpstring("method AddToolButton")] HRESULT AddToolButton([in] BSTR bstrButtonID, [in] BSTR bstrButtonText, [in] BSTR bstrIconFileName, [in] BSTR bstrClickEventFunction, [in, optional] VARIANT vOptions);
    [id(11), helpstring("method RemoveToolButton")] HRESULT RemoveToolButton([in] BSTR bstrButtonID);
    [id(12), helpstring("method SetToolButtonState")] HRESULT SetToolButtonState([in] BSTR bstrButtonID, [in] BSTR bstrState);
    [id(13), helpstring("method GetToolButtonRect")] HRESULT GetToolButtonRect([in] BSTR bstrButtonID, [out,retval] BSTR* pbstrRect);
    [id(14), helpstring("method GetPluginPath")] HRESULT GetPluginPath([in] BSTR bstrScriptFileName, [out,retval] BSTR* pbstrPluginPath);
    [id(15), helpstring("method AddScript")] HRESULT AddScript([in] BSTR bstrScriptFileName);
    [id(16), helpstring("method AddPluginMenu")] HRESULT AddPluginMenu([in] BSTR bstrMenuID, [in] BSTR bstrMenuText, [in] BSTR bstrClickEventFunction);
    [propget, id(17), helpstring("property SettingsPath")] HRESULT SettingsPath([out, retval] BSTR* pVal);
};
[
    object,
    uuid(A062E5F6-66ED-47E1-A7A9-A5417B96E510),
    dual,
    nonextensible,
    helpstring("IWizCommonTools Interface"),
    pointer_default(unique)
]
interface IWizCommonTools : IDispatch{
    [id(1), helpstring("method ExtractFilePath")] HRESULT ExtractFilePath([in] BSTR bstrFullPathName, [out,retval] BSTR* pbstrPath);
    [id(2), helpstring("method ExractFileName")] HRESULT ExractFileName([in] BSTR bstrFullPathName, [out,retval] BSTR* pbstrFileName);
    [id(3), helpstring("method ExtractFileExt")] HRESULT ExtractFileExt([in] BSTR bstrFileName, [out,retval] BSTR* pbstrExt);
    [id(4), helpstring("method LoadTextFromFile")] HRESULT LoadTextFromFile([in] BSTR bstrFileName, [out,retval] BSTR* pbstrText);
    [id(5), helpstring("method CreateActiveXObject")] HRESULT CreateActiveXObject([in] BSTR bstrProgID, [out,retval] IDispatch** ppObjectDisp);
    [id(6), helpstring("method InputBox")] HRESULT InputBox([in] BSTR bstrTitle, [in] BSTR bstrDescription, [out,retval] BSTR* pbstrValue);
    [id(7), helpstring("method LoadString")] HRESULT LoadString([in] BSTR bstrFileName, [in] BSTR bstrStringName, [out,retval] BSTR* pbstrValue);
    [id(8), helpstring("method GetValueFromIni")] HRESULT GetValueFromIni([in] BSTR bstrFileName, [in] BSTR bstrSection, [in] BSTR bstrKey, [out,retval] BSTR* pbstrValue);
    [id(9), helpstring("method SetValueToIni")] HRESULT SetValueToIni([in] BSTR bstrFileName, [in] BSTR bstrSection, [in] BSTR bstrKey, [in] BSTR bstrValue);
    [id(10), helpstring("method Html2Mime")] HRESULT Html2Mime([in] BSTR bstrHtmlFileName, [in] BSTR bstrCustomHeader, [in] BSTR bstrAttachmentFileNames, [in] BSTR bstrMimeFileName);
    [id(11), helpstring("method GetValueFromReg")] HRESULT GetValueFromReg([in] BSTR bstrRoot, [in] BSTR bstrKey, [in] BSTR bstrValueName, [out,retval] BSTR* pbstrData);
    [id(12), helpstring("method SetValueToReg")] HRESULT SetValueToReg([in] BSTR bstrRoot, [in] BSTR bstrKey, [in] BSTR bstrValueName, [in] BSTR bstrData, [in] BSTR bstrDataType);
    [id(13), helpstring("method EnumRegValues")] HRESULT EnumRegValues([in] BSTR bstrRoot, [in] BSTR bstrKey, [out, retval] VARIANT* pvValueNames);
    [id(14), helpstring("method EnumRegKeys")] HRESULT EnumRegKeys([in] BSTR bstrRoot, [in] BSTR bstrKey, [out, retval] VARIANT* pvKeys);
};
[
    object,
    uuid(EAC382F6-AD30-44BF-8853-2892372DC005),
    dual,
    nonextensible,
    helpstring("IWizDbxFile Interface"),
    pointer_default(unique)
]
interface IWizDbxFile : IDispatch{
    [id(1), helpstring("method Open")] HRESULT Open([in] BSTR bstrDbxFileName);
    [id(2), helpstring("method Close")] HRESULT Close(void);
    [propget, id(3), helpstring("property MailCount")] HRESULT MailCount([out, retval] LONG* pVal);
    [propget, id(4), helpstring("property MailItem")] HRESULT MailItem([in] LONG nIndex, [out, retval] IDispatch** pVal);
};
[
    object,
    uuid(0E8614A8-99D5-4BF7-AB9B-E5A9741FF7CC),
    dual,
    nonextensible,
    helpstring("IWizDbxMail Interface"),
    pointer_default(unique)
]
interface IWizDbxMail : IDispatch{
    [propget, id(1), helpstring("property MailProperty")] HRESULT MailProperty([in] BSTR bstrPropertyName, [out, retval] BSTR* pVal);
};
[
    uuid(3C7D5921-EBBE-4597-8DCC-7FB4F4DB228F),
    version(1.0),
    helpstring("WizTools 1.0 Type Library")
]
library WizToolsLib
{
    importlib("stdole2.tlb");

    [
        uuid(45E92139-A95B-401C-88FB-842EE284D170),
        helpstring("WizHtmlEditorApp Class")
    ]
    coclass WizHtmlEditorApp
    {
        [default] interface IWizHtmlEditorApp;
    };
    [
        uuid(4511852F-A672-426F-8860-B0B688D2A467),
        helpstring("WizCommonTools Class")
    ]
    coclass WizCommonTools
    {
        [default] interface IWizCommonTools;
    };
    [
        uuid(80DA3F56-12C9-4588-A0C1-A67C51EED025),
        helpstring("WizDbxFile Class")
    ]
    coclass WizDbxFile
    {
        [default] interface IWizDbxFile;
    };
    [
        uuid(F0864DA7-C573-4FC1-92EC-2A36A518BBA3),
        helpstring("WizDbxMail Class")
    ]
    coclass WizDbxMail
    {
        [default] interface IWizDbxMail;
    };
};
```
