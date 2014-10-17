```
// WizKMControls.idl : IDL source for WizKMControls
//

// This file will be processed by the MIDL tool to
// produce the type library (WizKMControls.tlb) and marshalling code.

#include "olectl.h"
#include "olectl.h"
#include "olectl.h"
#include "olectl.h"
import "oaidl.idl";
import "ocidl.idl";


[
    object,
    uuid(C128ECE0-A006-4E57-8054-4CBC49818231),
    dual,
    nonextensible,
    helpstring("IWizDocumentListCtrl Interface"),
    pointer_default(unique)
]interface IWizDocumentListCtrl : IDispatch{
    [propget, id(1), helpstring("property EventsListener")] HRESULT EventsListener([out, retval] VARIANT* pVal);
    [propput, id(1), helpstring("property EventsListener")] HRESULT EventsListener([in] VARIANT newVal);
    [propget, id(2), helpstring("property App")] HRESULT App([out, retval] IDispatch** pVal);
    [propput, id(2), helpstring("property App")] HRESULT App([in] IDispatch* newVal);
    [propget, id(3), helpstring("property Database")] HRESULT Database([out, retval] IDispatch** pVal);
    [propput, id(3), helpstring("property Database")] HRESULT Database([in] IDispatch* newVal);
    [propget, id(4), helpstring("property SelectedDocuments")] HRESULT SelectedDocuments([out, retval] IDispatch** pVal);
    [propput, id(4), helpstring("property SelectedDocuments")] HRESULT SelectedDocuments([in] IDispatch* newVal);
    [propget, id(5), helpstring("property StateSection")] HRESULT StateSection([out, retval] BSTR* pVal);
    [propput, id(5), helpstring("property StateSection")] HRESULT StateSection([in] BSTR newVal);
    [propget, id(6), helpstring("property Border")] HRESULT Border([out, retval] VARIANT_BOOL* pVal);
    [propput, id(6), helpstring("property Border")] HRESULT Border([in] VARIANT_BOOL newVal);
    [propget, id(7), helpstring("property DisableContextMenu")] HRESULT DisableContextMenu([out, retval] VARIANT_BOOL* pVal);
    [propput, id(7), helpstring("property DisableContextMenu")] HRESULT DisableContextMenu([in] VARIANT_BOOL newVal);
    [propget, id(8), helpstring("property SortBy")] HRESULT SortBy([out, retval] BSTR* pVal);
    [propput, id(8), helpstring("property SortBy")] HRESULT SortBy([in] BSTR newVal);
    [propget, id(9), helpstring("property SortOrder")] HRESULT SortOrder([out, retval] BSTR* pVal);
    [propput, id(9), helpstring("property SortOrder")] HRESULT SortOrder([in] BSTR newVal);
    [propget, id(10), helpstring("property ParentFolder")] HRESULT ParentFolder([out, retval] IDispatch** pVal);
    [id(11), helpstring("method SetDocuments")] HRESULT SetDocuments([in] IDispatch* pDisp);
    [id(12), helpstring("method SetDocuments2")] HRESULT SetDocuments2([in] IDispatch* pDisp, [in] BSTR bstrSortBy, [in] BSTR bstrSortOrder);
    [id(13), helpstring("method Refresh")] HRESULT Refresh();
    [id(14), helpstring("method GetDocuments")] HRESULT GetDocuments([out, retval] IDispatch** pVal);
    [propget, id(15), helpstring("property Options")] HRESULT Options([out, retval] LONG* pVal);
    [propput, id(15), helpstring("property Options")] HRESULT Options([in] LONG newVal);
    [propget, id(16), helpstring("property SecondLineType")] HRESULT SecondLineType([out, retval] LONG* pVal);
    [propput, id(16), helpstring("property SecondLineType")] HRESULT SecondLineType([in] LONG newVal);
    [id(17), helpstring("method SetFilter")] HRESULT SetFilter([in] IUnknown* pFilterUnk);
    [id(18), helpstring("method RefreshDocuments")] HRESULT RefreshDocuments();
    [propget, id(19), helpstring("property ViewType")] HRESULT ViewType([out, retval] LONG* pVal);
    [propput, id(19), helpstring("property ViewType")] HRESULT ViewType([in] LONG newVal);
};

[
    object,
    uuid(39B2717D-7FDA-4EDD-91A4-0173FD35B871),
    dual,
    nonextensible,
    helpstring("IWizDocumentAttachmentListCtrl Interface"),
    pointer_default(unique)
]
interface IWizDocumentAttachmentListCtrl : IDispatch{
    [propget, id(1), helpstring("property EventsListener")] HRESULT EventsListener([out, retval] VARIANT* pVal);
    [propput, id(1), helpstring("property EventsListener")] HRESULT EventsListener([in] VARIANT newVal);
    [propget, id(2), helpstring("property App")] HRESULT App([out, retval] IDispatch** pVal);
    [propput, id(2), helpstring("property App")] HRESULT App([in] IDispatch* newVal);
    [propget, id(3), helpstring("property Database")] HRESULT Database([out, retval] IDispatch** pVal);
    [propput, id(3), helpstring("property Database")] HRESULT Database([in] IDispatch* newVal);
    [propget, id(4), helpstring("property Document")] HRESULT Document([out, retval] IDispatch** pVal);
    [propput, id(4), helpstring("property Document")] HRESULT Document([in] IDispatch* newVal);
    [propget, id(5), helpstring("property SelectedAttachments")] HRESULT SelectedAttachments([out, retval] IDispatch** pVal);
    [propput, id(5), helpstring("property SelectedAttachments")] HRESULT SelectedAttachments([in] IDispatch* newVal);
    [propget, id(6), helpstring("property Border")] HRESULT Border([out, retval] VARIANT_BOOL* pVal);
    [propput, id(6), helpstring("property Border")] HRESULT Border([in] VARIANT_BOOL newVal);
    [id(7), helpstring("method AddAttachments")] HRESULT AddAttachments(void);
    [id(8), helpstring("method AddAttachments2")] HRESULT AddAttachments2([in] VARIANT* pvFiles);
    [id(9), helpstring("method GetMinSize")] HRESULT GetMinSize([out, retval] ULONG* pnsize);
};

[
    object,
    uuid(B671315F-B6FE-4799-AE50-AEA0848FBA05),
    dual,
    nonextensible,
    helpstring("IWizCommonUI Interface"),
    pointer_default(unique)
]
interface IWizCommonUI : IDispatch{
    [id(1), helpstring("method QueryUserAccount")] HRESULT QueryUserAccount([in] LONGLONG hWnd, [in, out] BSTR* pbstrUserName, [in, out] BSTR* pbstrPassword);
    [id(2), helpstring("method AddDocumentToCalendar")] HRESULT AddDocumentToCalendar([in] IDispatch* pDocumentDisp, [out,retval] VARIANT_BOOL* pbRet);
    [id(3), helpstring("method CreateCalendarEvent")] HRESULT CreateCalendarEvent([in] IDispatch* pDatabaseDisp, [in] DATE dtEvent, [out,retval] IDispatch** ppDocumentDisp);
    [id(4), helpstring("method EditCalendarEvent")] HRESULT EditCalendarEvent([in] IDispatch* pDocumentDisp, [out,retval] VARIANT_BOOL* pbRet);
    [id(5), helpstring("method AboutBox")] HRESULT AboutBox([in] BSTR bstrModuleName);
    [id(6), helpstring("method OptionsDlg")] HRESULT OptionsDlg([in] LONG nFlags);
    [id(7), helpstring("method BackupDlg")] HRESULT BackupDlg([in] LONG nFlags);
    [id(8), helpstring("method AccountsManagerDlg")] HRESULT AccountsManagerDlg([in] LONG nFlags);
    [id(9), helpstring("method ChooseAccount")] HRESULT ChooseAccount([in] LONG nFlags, [out,retval] BSTR* pbstrDatabasePath);
    [id(10), helpstring("method GetIntValue")] HRESULT GetIntValue([in] BSTR bstrTitle, [in] BSTR bstrDescription, [in] LONG nInitValue, [out] VARIANT_BOOL* pvbOK, [out,retval] LONG* pnRet);
    [id(11), helpstring("method GetIntValue2")] HRESULT GetIntValue2([in] BSTR bstrTitle, [in] BSTR bstrDescription, [in] LONG nInitValue, [in] LONG nMin, [in] LONG nMax, [out] VARIANT_BOOL* pvbOK, [out,retval] LONG* pnRet);
    [id(12), helpstring("method CheckNewVersion")] HRESULT CheckNewVersion([in] VARIANT_BOOL vbPrompt, [in] VARIANT_BOOL vbSilent);
    [id(13), helpstring("method InputBox")] HRESULT InputBox([in] BSTR bstrTitle, [in] BSTR bstrDescription, [in] BSTR bstrInitValue, [out,retval] BSTR* pbstrText);
    [id(14), helpstring("method InputMultiLineText")] HRESULT InputMultiLineText([in] BSTR bstrTitle, [in] BSTR bstrDescription, [in] BSTR bstrInitValue, [out,retval] BSTR* pbstrText);
    [id(15), helpstring("method SelectWindowsFile")] HRESULT SelectWindowsFile([in] VARIANT_BOOL vbOpen, [in] BSTR bstrFilter, [out,retval] BSTR* pbstrResultFileName);
    [id(16), helpstring("method SelectWindowsFolder")] HRESULT SelectWindowsFolder([in] BSTR bstrDescription, [out,retval] BSTR* pbtrFolderPath);
    [id(17), helpstring("method SelectWizFolder")] HRESULT SelectWizFolder([in] BSTR bstrDatabasePath, [in] BSTR bstrDescription, [out,retval] IDispatch** ppFolderDisp);
    [id(18), helpstring("method SelectWizDocument")] HRESULT SelectWizDocument([in] BSTR bstrDatabasePath, [in] BSTR bstrDescription, [out,retval] IDispatch** ppDocumentDisp);
    [id(19), helpstring("method LoadTextFromFile")] HRESULT LoadTextFromFile([in] BSTR bstrFileName, [out,retval] BSTR* pbstrText);
    [id(20), helpstring("method SaveTextToFile")] HRESULT SaveTextToFile([in] BSTR bstrFileName, [in] BSTR bstrText, [in] BSTR bstrCharset);
    [id(21), helpstring("method GetValueFromIni")] HRESULT GetValueFromIni([in] BSTR bstrFileName, [in] BSTR bstrSection, [in] BSTR bstrKey, [out,retval] BSTR* pbstrValue);
    [id(22), helpstring("method SetValueToIni")] HRESULT SetValueToIni([in] BSTR bstrFileName, [in] BSTR bstrSection, [in] BSTR bstrKey, [in] BSTR bstrValue);
    [id(23), helpstring("method GetATempFileName")] HRESULT GetATempFileName([in] BSTR bstrFileExt, [out,retval] BSTR* pbstrFileName);
    [id(24), helpstring("method GetSpecialFolder")] HRESULT GetSpecialFolder([in] BSTR bstrFolderName, [out,retval] BSTR* pbstrPath);
    [id(26), helpstring("method EnumFiles")] HRESULT EnumFiles([in] BSTR bstrPath, [in] BSTR bstrFileExt, [in] VARIANT_BOOL vbIncludeSubFolders, [out,retval] VARIANT* pvFiles);
    [id(27), helpstring("method EnumFolders")] HRESULT EnumFolders([in] BSTR bstrPath, [out,retval] VARIANT* pvFolders);
    [id(28), helpstring("method CreateActiveXObject")] HRESULT CreateActiveXObject([in] BSTR bstrProgID, [out,retval] IDispatch** ppObjDisp);
    [id(29), helpstring("method GetValueFromReg")] HRESULT GetValueFromReg([in] BSTR bstrRoot, [in] BSTR bstrKey, [in] BSTR bstrValueName, [out,retval] BSTR* pbstrData);
    [id(30), helpstring("method SetValueToReg")] HRESULT SetValueToReg([in] BSTR bstrRoot, [in] BSTR bstrKey, [in] BSTR bstrValueName, [in] BSTR bstrData, [in] BSTR bstrDataType);
    [id(31), helpstring("method EnumRegValues")] HRESULT EnumRegValues([in] BSTR bstrRoot, [in] BSTR bstrKey, [out, retval] VARIANT* pvValueNames);
    [id(32), helpstring("method EnumRegKeys")] HRESULT EnumRegKeys([in] BSTR bstrRoot, [in] BSTR bstrKey, [out, retval] VARIANT* pvKeys);
    [id(33), helpstring("method SelectTemplate")] HRESULT SelectTemplate([in] BSTR bstrTemplateType, [out, retval] BSTR* pbstrResultTemplateFileName);
    [id(34), helpstring("method HtmlTagGetAttributeValue")] HRESULT HtmlTagGetAttributeValue([in] BSTR bstrHtmlTag, [in] BSTR bstrTagAttributeName, [out, retval] BSTR* pbstrAttributeValue);
    [id(35), helpstring("method HtmlExtractTags")] HRESULT HtmlExtractTags([in] BSTR bstrHtmlText, [in] BSTR bstrTagName, [in] BSTR bstrTagAttributeName, [in] BSTR bstrTagAttributeValue, [out, retval] VARIANT* pvTags);
    [id(36), helpstring("method HtmlEnumLinks")] HRESULT HtmlEnumLinks([in] BSTR bstrHtmlText, [in] BSTR bstrURL, [out, retval] VARIANT* pvLinks);
    [id(37), helpstring("method HtmlGetContent")] HRESULT HtmlGetContent([in] BSTR bstrHtmlText, [in] BSTR bstrURL, [out, retval] BSTR* pbstrContent);
    [id(38), helpstring("method HtmlConvertFileToHtmlFile")] HRESULT HtmlConvertFileToHtmlFile([in] BSTR bstrSrcFileName, [out, retval] BSTR* pbstrHtmlFileName);
    [id(39), helpstring("method HtmlConvertZipFileToHtmlFile")] HRESULT HtmlConvertZipFileToHtmlFile([in] BSTR bstrSrcZipFileName, [in] BSTR bstrDestHtmlFileName, [in] BSTR bstrHtmlTitle);
    [id(40), helpstring("method OptionsDlg2")] HRESULT OptionsDlg2([in] IDispatch* pDisp, [in] LONG nFlags);
    [id(41), helpstring("method SelectDay")] HRESULT SelectDay([in] DATE dtDefault, [out, retval] DATE* pRet);
    [id(42), helpstring("method ReportSyncResult")] HRESULT ReportSyncResult([in] LONGLONG hWnd, [in] VARIANT vDownloaded, [in] VARIANT vUploaded, [in] VARIANT vSkipped, [in] VARIANT vFailed);
    [id(43), helpstring("method CopyTextToClipboard")] HRESULT CopyTextToClipboard([in] BSTR bstrText);
    [id(44), helpstring("method PromptTrafficLimit")] HRESULT PromptTrafficLimit([in] IDispatch* pDatabaseDisp, [in] LONGLONG hWnd, [in] LONGLONG nLimit);
    [id(45), helpstring("method AccountSecurityDlg")] HRESULT AccountSecurityDlg([in] IDispatch* pDisp, [in] LONG nFlags);
    [id(46), helpstring("method InitEncryptionDlg")] HRESULT InitEncryptionDlg([in] IDispatch* pDisp, [in] LONG nFlags);
    [id(47), helpstring("method ImportCertDlg")] HRESULT ImportCertDlg([in] IDispatch* pDisp, [in] LONG nFlags);
    [id(48), helpstring("method EnterPasswordDlg")] HRESULT EnterPasswordDlg([in] BSTR bstrTitle, [in] BSTR bstrInfo, [in] VARIANT_BOOL vbShowRememberPasswordButton, [out] VARIANT_BOOL* pvbRememberPassword, [out, retval] BSTR* pbstrPassword);
    [id(49), helpstring("method QueryUserAccount2")] HRESULT QueryUserAccount2([in] IDispatch* pDatabaseDisp, [in] LONGLONG hWnd, [in, out] BSTR* pbstrUserName, [in, out] BSTR* pbstrPassword);
    [id(50), helpstring("method ClipboardToImage")] HRESULT ClipboardToImage([in] OLE_HANDLE hwnd, [in] BSTR bstrOptions, [out,retval] BSTR* pbstrImageFileName);
    [id(51), helpstring("method CaptureScreen")] HRESULT CaptureScreen([in] BSTR bstrOptions, [out,retval] BSTR* pbstrImageFileName);
    [id(52), helpstring("method DumpCHM")] HRESULT DumpCHM([in] BSTR bstrCHMFileName, [in] BSTR bstrTempPath, [in] BSTR bstrOptions, [out,retval] VARIANT* pvRet);
    [id(53), helpstring("method HyperlinkDlg")] HRESULT HyperlinkDlg([in] IDispatch* pdispDatabase, [in] IDispatch* pdispDocument, [in] IDispatch* pdispHtmlDocument, [in] BSTR bstrInitURL, [out,retval] BSTR* pbstrResultURL);
    [id(54), helpstring("method SelectColor")] HRESULT SelectColor([in] BSTR bstrInitColor, [out,retval] BSTR* pbstrRetColor);
    [id(55), helpstring("method GetFontNames")] HRESULT GetFontNames([out,retval] VARIANT* pvNames);
    [id(56), helpstring("method SelectFont")] HRESULT SelectFont([in] BSTR bstrInit, [out,retval] BSTR* pbstrRet);
    [id(57), helpstring("method CreateWizObject")] HRESULT CreateWizObject([in] BSTR bstrObjectName, [out,retval] IDispatch** ppVal);
    [id(58), helpstring("method GetScriptParamValue")] HRESULT GetScriptParamValue([in] BSTR bstrParamName, [out,retval] BSTR* pbstrVal);
    [id(59), helpstring("method LoadStringFromFile")] HRESULT LoadStringFromFile([in] BSTR bstrFileName, [in] BSTR bstrStringName, [out, retval] BSTR* pVal);
    [id(60), helpstring("method URLDownloadToText")] HRESULT URLDownloadToText([in] BSTR bstrURL, [out, retval] BSTR* pbstrText);
    [id(61), helpstring("method SelectWizTags")] HRESULT SelectWizTags([in] BSTR bstrDatabasePath, [in] BSTR bstrTitle, [in] LONG nFlags, [out,retval] IDispatch** ppTagCollectionDisp);
    [id(62), helpstring("method CalendarDlg")] HRESULT CalendarDlg([in] LONG nFlags);
    [id(63), helpstring("method AccountInfoDlg")] HRESULT AccountInfoDlg([in] IDispatch* pDatabaseDisp, [in] LONG nFlags);
    [id(64), helpstring("method GetIniSections")] HRESULT GetIniSections([in] BSTR bstrFileName, [out, retval] VARIANT* pvValueNames);
    [id(65), helpstring("method LoadStringFromFile2")] HRESULT LoadStringFromFile2([in] BSTR bstrFileName, [in] BSTR bstrSectionName, [in] BSTR bstrStringName, [out, retval] BSTR* pVal);
    [id(66), helpstring("method ShowHtmlDialog")] HRESULT ShowHtmlDialog([in] BSTR bstrTitle, [in] BSTR bstrURL, [in] LONG nWidth, [in] LONG nHeight, [in] BSTR bstrExtOptions, [in] VARIANT vParam, [out,retval] VARIANT* pvRet);
    [id(67), helpstring("method CloseHtmlDialog")] HRESULT CloseHtmlDialog([in] IDispatch* pdispHtmlDialogDocument, [in] VARIANT vRet);
    [id(68), helpstring("method ShowMessage")] HRESULT ShowMessage([in] BSTR bstrText, [in] BSTR bstrTitle, [in] LONG nType, [out,retval] LONG* pnRet);
    [id(69), helpstring("method CombineURL")] HRESULT CombineURL([in] BSTR bstrBase, [in] BSTR bstrPart, [out,retval] BSTR* pbstrResult);
    [id(70), helpstring("method URLDownloadToTempFile")] HRESULT URLDownloadToTempFile([in] BSTR bstrURL, [out,retval] BSTR* pbstrTempFileName);
    [id(71), helpstring("method EditDocument")] HRESULT EditDocument([in] IDispatch* pApp, [in] IDispatch* pEvents, [in] IDispatch* pDocumentDisp, [in] BSTR bstrOptions, [out,retval] LONGLONG* pnWindowHandle);
    [id(72), helpstring("method NewDocument")] HRESULT NewDocument([in] IDispatch* pApp, [in] IDispatch* pEvents, [in] IDispatch* pFolderDisp, [in] BSTR bstrOptions, [out,retval] LONGLONG* pnWindowHandle);
    [id(73), helpstring("method ClearDocumentWindow")] HRESULT ClearDocumentWindow();
    [id(74), helpstring("method CreateAccountDlg")] HRESULT CreateAccountDlg([in] BSTR bstrOptions, [out,retval] BSTR* pbstrRetDatabasePath);
    [id(75), helpstring("method AddExistingAccountDlg")] HRESULT AddExistingAccountDlg([in] BSTR bstrOptions, [out,retval] BSTR* pbstrRetDatabasePath);
    [id(76), helpstring("method ChangePasswordDlg")] HRESULT ChangePasswordDlg([in] IDispatch* pDatabaseDisp);
    [id(77), helpstring("method ProtectAccountDlg")] HRESULT ProtectAccountDlg([in] IDispatch* pDatabaseDisp);
    [id(78), helpstring("method RefreshDatabaseIndex")] HRESULT RefreshDatabaseIndex([in] LONGLONG nWindowHandle, [in] IDispatch* pDatabaseDisp);
    [id(79), helpstring("method ShareDocument")] HRESULT ShareDocument([in] LONGLONG nWindowHandle, [in] IDispatch* pDocumentDisp, [in] BSTR bstrParams);
    [id(80), helpstring("method ShareDocumentDlg")] HRESULT ShareDocumentDlg([in] IDispatch* pDocumentDisp, [in] BSTR bstrParams);
    [id(81), helpstring("method RefreshAccountInfo")] HRESULT RefreshAccountInfo([in] IDispatch* pDatabaseDisp);
    [id(82), helpstring("method PathFileExists")] HRESULT PathFileExists([in] BSTR bstrPath, [out,retval] VARIANT_BOOL* pvbExists);
    [id(83), helpstring("method ExtractFilePath")] HRESULT ExtractFilePath([in] BSTR bstrFileName, [out,retval] BSTR* pbstrFilePath);
    [id(84), helpstring("method ExtractFileName")] HRESULT ExtractFileName([in] BSTR bstrFileName, [out,retval] BSTR* pbstrName);
    [id(85), helpstring("method ExtractFileExt")] HRESULT ExtractFileExt([in] BSTR bstrFileName, [out,retval] BSTR* pbstrExt);
    [id(86), helpstring("method ExtractFileTitle")] HRESULT ExtractFileTitle([in] BSTR bstrFileName, [out,retval] BSTR* pbstrTitle);
    [id(88), helpstring("method CreateDirectory")] HRESULT CreateDirectory([in] BSTR bstrPath);
    [id(89), helpstring("method CopyFile")] HRESULT CopyFile([in] BSTR bstrExistingFile, [in] BSTR bstrNewFileName);
    [id(90), helpstring("method EnumFiles2")] HRESULT EnumFiles2([in] BSTR bstrPath, [in] BSTR bstrFileExt, [in] VARIANT_BOOL vbIncludeSubFolders, [out,retval] BSTR* pvFiles);
    [id(91), helpstring("method EnumFolders2")] HRESULT EnumFolders2([in] BSTR bstrPath, [out,retval] BSTR* pvFolders);
    [id(92), helpstring("method EnumRegValues2")] HRESULT EnumRegValues2([in] BSTR bstrRoot, [in] BSTR bstrKey, [out, retval] BSTR* pvValueNames);
    [id(93), helpstring("method EnumRegKeys2")] HRESULT EnumRegKeys2([in] BSTR bstrRoot, [in] BSTR bstrKey, [out, retval] BSTR* pvKeys);
    [id(94), helpstring("method HtmlExtractTags2")] HRESULT HtmlExtractTags2([in] BSTR bstrHtmlText, [in] BSTR bstrTagName, [in] BSTR bstrTagAttributeName, [in] BSTR bstrTagAttributeValue, [out, retval] BSTR* pvTags);
    [id(95), helpstring("method HtmlEnumLinks2")] HRESULT HtmlEnumLinks2([in] BSTR bstrHtmlText, [in] BSTR bstrURL, [out, retval] BSTR* pvLinks);
    [id(96), helpstring("method GetFontNames2")] HRESULT GetFontNames2([out,retval] BSTR* pvNames);
VARIANT_BOOL vbWait, [out,retval] long* pnRet);
    [id(98), helpstring("method GetLongPathName")] HRESULT GetLongPathName([in] BSTR bstrPath, [out, retval] BSTR* pbstrRet);
    [id(99), helpstring("method GetLongPathName")] HRESULT GetShortPathName([in] BSTR bstrPath, [out, retval] BSTR* pbstrRet);
    [id(100), helpstring("method SelectWizFolder2")] HRESULT SelectWizFolder2([in] BSTR bstrDatabasePath, [in] BSTR bstrPassword, [in] BSTR bstrDescription, [out,retval] IDispatch** ppFolderDisp);
    [id(101), helpstring("method UserIdToDatabasePath")] HRESULT UserIdToDatabasePath([in] BSTR bstrUserId, [out,retval] BSTR* ppbstrDatabasePath);
    [id(102), helpstring("method InputBox2")] HRESULT InputBox2([in] OLE_HANDLE hwndParent, [in] VARIANT_BOOL vbPassword, [in] BSTR bstrIconFileName, [in] BSTR bstrTitle, [in] BSTR bstrDescription, [in] BSTR bstrInitValue, [out,retval] BSTR* pbstrText);
    [id(103), helpstring("method ShareDocument2")] HRESULT ShareDocument2([in] IDispatch* pApp, [in] IDispatch* pDocumentDisp, [in] BSTR bstrParams);
    [id(104), helpstring("method GetFileSize")] HRESULT GetFileSize([in] BSTR bstrFileName, [out, retval] LONGLONG* pnSize);
    [id(105), helpstring("method IntToByteSizeString")] HRESULT IntToByteSizeString([in] LONG nSize, [out, retval] BSTR* pbstr);
    [id(106), helpstring("method ShowLastInputWindow")] HRESULT ShowLastInputWindow();
    [id(107), helpstring("method AccountInfoDlg2")] HRESULT AccountInfoDlg2([in] IDispatch* pApp, [in] IDispatch* pDatabaseDisp, [in] LONG nFlags);
};
[
    object,
    uuid(49C2236A-6DA1-4AA9-BAD7-BF0DA94535A4),
    dual,
    nonextensible,
    helpstring("IWizBatchDownloader Interface"),
    pointer_default(unique)
]
interface IWizBatchDownloader : IDispatch{
    [id(1), helpstring("method AddJob")] HRESULT AddJob([in] BSTR bstrDatabasePath, [in] BSTR bstrLocation, [in] BSTR bstrURL, [in] BSTR bstrLinkText, [in] LONG nUpdateDocumentFlags, [in] VARIANT_BOOL vbLinkTextAsTitle, [in] VARIANT_BOOL vbExecuteScript);
    [id(2), helpstring("method AddJobsFromTextFile")] HRESULT AddJobsFromTextFile([in] BSTR bstrDatabasePath, [in] BSTR bstrLocation, [in] BSTR bstrTextFileName, [in] LONG nUpdateDocumentFlags, [in] VARIANT_BOOL vbExecuteScript);
    [id(3), helpstring("method ShowWindow")] HRESULT ShowWindow(VARIANT_BOOL vbStartDownload);
    [propget, id(4), helpstring("property CombineDocuments")] HRESULT CombineDocuments([out, retval] VARIANT_BOOL* pVal);
    [propput, id(4), helpstring("property CombineDocuments")] HRESULT CombineDocuments([in] VARIANT_BOOL newVal);
    [id(5), helpstring("method AddJob2")] HRESULT AddJob2([in] BSTR bstrDatabasePath, [in] BSTR bstrLocation, [in] BSTR bstrURL, [in] BSTR bstrLinkText, [in] LONG nUpdateDocumentFlags, [in] LONG nJobFlags);
    [id(6), helpstring("method AddJob3")] HRESULT AddJob3([in] BSTR bstrDatabasePath, [in] BSTR bstrOverwriteDocumentGUID, [in] BSTR bstrURL, [in] BSTR bstrTitle, [in] LONG nUpdateDocumentFlags, [in] LONG nJobFlags);
    [propget, id(7), helpstring("property Title")] HRESULT Title([out, retval] BSTR* pVal);
    [propput, id(7), helpstring("property Title")] HRESULT Title([in] BSTR newVal);
    [id(8), helpstring("method SaveJobsToFile")] HRESULT SaveJobsToFile([in] BSTR bstrFileName);
    [id(9), helpstring("method LoadJobsFromFile")] HRESULT LoadJobsFromFile([in] BSTR bstrFileName);
    [id(10), helpstring("method AddJob")] HRESULT AddJob4([in] BSTR bstrDatabasePath, [in] BSTR bstrLocation, [in] BSTR bstrURL, [in] BSTR bstrLinkText, [in] LONG nUpdateDocumentFlags, [in] VARIANT_BOOL vbLinkTextAsTitle, [in] VARIANT_BOOL vbExecuteScript, [in] BSTR bstrKbGUID, [in] BSTR bstrTagGUIDs);
};
[
    object,
    uuid(FF5BCE34-5FC1-4337-80D3-87747643CB15),
    dual,
    nonextensible,
    helpstring("IWizProgressWindow Interface"),
    pointer_default(unique)
]
interface IWizProgressWindow : IDispatch{
    [propget, id(1), helpstring("property Title")] HRESULT Title([out, retval] BSTR* pVal);
    [propput, id(1), helpstring("property Title")] HRESULT Title([in] BSTR newVal);
    [propget, id(2), helpstring("property Max")] HRESULT Max([out, retval] LONG* pVal);
    [propput, id(2), helpstring("property Max")] HRESULT Max([in] LONG newVal);
    [propget, id(3), helpstring("property Pos")] HRESULT Pos([out, retval] LONG* pVal);
    [propput, id(3), helpstring("property Pos")] HRESULT Pos([in] LONG newVal);
    [propget, id(4), helpstring("property Text")] HRESULT Text([out, retval] BSTR* pVal);
    [propput, id(4), helpstring("property Text")] HRESULT Text([in] BSTR newVal);
    [id(5), helpstring("method Show")] HRESULT Show(void);
    [id(6), helpstring("method Hide")] HRESULT Hide(void);
    [id(7), helpstring("method Destroy")] HRESULT Destroy(void);
    [id(8), helpstring("method SetParent")] HRESULT SetParent([in] OLE_HANDLE hwnd);
};
[
    object,
    uuid(476F255C-8032-42A4-AA36-4EBB25C36BAB),
    dual,
    nonextensible,
    helpstring("IWizSyncProgressDlg Interface"),
    pointer_default(unique)
]
interface IWizSyncProgressDlg : IDispatch{
    [id(1), helpstring("method Show")] HRESULT Show(void);
    [id(2), helpstring("method Hide")] HRESULT Hide(void);
    [id(3), helpstring("method SetDatabasePath")] HRESULT SetDatabasePath([in] BSTR bstrDatabasePath);
    [id(4), helpstring("method AutoShow")] HRESULT AutoShow(void);
    [id(5), helpstring("method Reset")] HRESULT Reset(void);
    [propget, id(6), helpstring("property Background")] HRESULT Background([out, retval] VARIANT_BOOL* pVal);
    [propput, id(6), helpstring("property Background")] HRESULT Background([in] VARIANT_BOOL newVal);
    [propget, id(7), helpstring("property Window")] HRESULT Window([out, retval] LONGLONG* pVal);

};
[
    object,
    uuid(9A7BFBD0-65CE-4F73-A819-BAE5819C1127),
    dual,
    nonextensible,
    helpstring("IWizHtmlDialogExternal Interface"),
    pointer_default(unique)
]
interface IWizHtmlDialogExternal : IDispatch{
    [propget, id(1), helpstring("property HtmlDialogParam")] HRESULT HtmlDialogParam([out, retval] VARIANT* pVal);
    [propget, id(2), helpstring("property CommonUIObject")] HRESULT CommonUIObject([out, retval] IDispatch** pVal);
};
[
    object,
    uuid(C163AE7B-E5B1-4789-A61B-27AFC8C9E17D),
    dual,
    nonextensible,
    helpstring("IWizCategoryCtrl Interface"),
    pointer_default(unique)
]
interface IWizCategoryCtrl : IDispatch{
    [propget, id(1), helpstring("property EventsListener")] HRESULT EventsListener([out, retval] VARIANT* pVal);
    [propput, id(1), helpstring("property EventsListener")] HRESULT EventsListener([in] VARIANT newVal);
    [propget, id(2), helpstring("property App")] HRESULT App([out, retval] IDispatch** pVal);
    [propput, id(2), helpstring("property App")] HRESULT App([in] IDispatch* newVal);
    [propget, id(3), helpstring("property Database")] HRESULT Database([out, retval] IDispatch** pVal);
    [propput, id(3), helpstring("property Database")] HRESULT Database([in] IDispatch* newVal);
    [propget, id(4), helpstring("property SelectedType")] HRESULT SelectedType([out, retval] LONG* pVal);
    [propget, id(5), helpstring("property SelectedFolder")] HRESULT SelectedFolder([out, retval] IDispatch** pVal);
    [propput, id(5), helpstring("property SelectedFolder")] HRESULT SelectedFolder([in] IDispatch* newVal);
    [propget, id(6), helpstring("property SelectedTags")] HRESULT SelectedTags([out, retval] IDispatch** pVal);
    [propput, id(6), helpstring("property SelectedTags")] HRESULT SelectedTags([in] IDispatch* newVal);
    [propget, id(7), helpstring("property SelectedStyle")] HRESULT SelectedStyle([out, retval] IDispatch** pVal);
    [propput, id(7), helpstring("property SelectedStyle")] HRESULT SelectedStyle([in] IDispatch* newVal);
    [propget, id(8), helpstring("property SelectedDocument")] HRESULT SelectedDocument([out, retval] IDispatch** pVal);
    [propput, id(8), helpstring("property SelectedDocument")] HRESULT SelectedDocument([in] IDispatch* newVal);
    [propget, id(9), helpstring("property StateSection")] HRESULT StateSection([out, retval] BSTR* pVal);
    [propput, id(9), helpstring("property StateSection")] HRESULT StateSection([in] BSTR newVal);
    [propget, id(10), helpstring("property Options")] HRESULT Options([out, retval] LONG* pVal);
    [propput, id(10), helpstring("property Options")] HRESULT Options([in] LONG newVal);
    [propget, id(11), helpstring("property ShowDocuments")] HRESULT ShowDocuments([out, retval] VARIANT_BOOL* pVal);
    [propput, id(11), helpstring("property ShowDocuments")] HRESULT ShowDocuments([in] VARIANT_BOOL newVal);
    [propget, id(12), helpstring("property Border")] HRESULT Border([out, retval] VARIANT_BOOL* pVal);
    [propput, id(12), helpstring("property Border")] HRESULT Border([in] VARIANT_BOOL newVal);
    [id(13), helpstring("method Refresh")] HRESULT Refresh([in] LONG nFlags);
    [id(14), helpstring("method ExecuteCommand")] HRESULT ExecuteCommand([in] BSTR bstrCommandName, [in] VARIANT* pvParam1, [in] VARIANT* pvParam2, [out, retval] VARIANT* pvRetValue);
    [id(15), helpstring("method GetSelectedDocuments2")] HRESULT GetSelectedDocuments2([out] BSTR* pbstrSortBy, [out] BSTR* pbstrSortOrder, [out, retval] IDispatch** pVal);
    [id(16), helpstring("method SaveState")] HRESULT SaveState();
    [id(17), helpstring("method GetSelectedItemCustomData")] HRESULT GetSelectedItemCustomData([out, retval] BSTR* pbstrData);
    [propget, id(18), helpstring("property CurrentDatabase")] HRESULT CurrentDatabase([out, retval] IDispatch** pVal);
    [id(19), helpstring("method SetSearchResult")] HRESULT SetSearchResult([in] IDispatch* pDatabaseDisp, [in] BSTR bstrKeywords, [in] IDispatch* pDocumentDisp, [in] VARIANT_BOOL vbForceSelect);
};

[
    object,
    uuid(140D38EC-AE4A-4AFE-94CD-FAEEEAB7D1B6),
    dual,
    nonextensible,
    helpstring("IWizStatusWindow Interface"),
    pointer_default(unique)
]
interface IWizStatusWindow : IDispatch{
    [propget, id(1), helpstring("property Text")] HRESULT Text([out, retval] BSTR* pVal);
    [propput, id(1), helpstring("property Text")] HRESULT Text([in] BSTR newVal);
    [id(2), helpstring("method Show")] HRESULT Show(void);
    [id(3), helpstring("method Hide")] HRESULT Hide(void);
};
[
    uuid(89FE0B90-C336-4BDF-A931-7CA8DCD7E002),
    version(1.0),
    helpstring("WizKMControls 1.0 Type Library")
]
library WizKMControlsLib
{
    importlib("stdole2.tlb");

    [
        uuid(D30F2928-D136-4F2D-98C4-080E3CB1C92C),
        control,
        helpstring("WizDocumentListCtrl Class")
    ]
    coclass WizDocumentListCtrl
    {
        [default] interface IWizDocumentListCtrl;
    };

    [
        uuid(A74098CA-0E98-40D6-92BF-06AAAE1B2EB8),
        control,
        helpstring("WizDocumentAttachmentListCtrl Class")
    ]
    coclass WizDocumentAttachmentListCtrl
    {
        [default] interface IWizDocumentAttachmentListCtrl;
    };
    [
        uuid(5EABDAD8-A056-4445-AC98-E66885B0935F),
        helpstring("WizCommonUI Class")
    ]
    coclass WizCommonUI
    {
        [default] interface IWizCommonUI;
    };
    [
        uuid(8C43A23A-11CD-4BFA-A3FA-CBC4A586F666),
        helpstring("WizBatchDownloader Class")
    ]
    coclass WizBatchDownloader
    {
        [default] interface IWizBatchDownloader;
    };

    enum WizCategorySelectedType
    { 
        [helpstring("None")]        wizCategorySelectedTypeNone =            0,
        [helpstring("Folder")]        wizCategorySelectedTypeFolder =            1,
        [helpstring("Document")]        wizCategorySelectedTypeDocument =    2,
        [helpstring("Tag")]        wizCategorySelectedTypeTag =                3,
        [helpstring("Style")]        wizCategorySelectedTypeStyle =            4,
        [helpstring("AllFolders")]        wizCategorySelectedTypeAllFolders = 5,
        [helpstring("AllTags")]        wizCategorySelectedTypeAllTags =        6,
        [helpstring("AllStyles")]        wizCategorySelectedTypeAllStyles =  7,
        [helpstring("AllQuickSearches")]        wizCategorySelectedTypeAllQuickSearches =  8,
        [helpstring("QuickSearchGroup")]        wizCategorySelectedTypeQuickSearchGroup =  9,
        [helpstring("QuickSearch")]        wizCategorySelectedTypeQuickSearch =  10,
        [helpstring("FavoriteFolder")]        wizCategorySelectedTypeFavoriteFolder =            11,
        [helpstring("AllFavoriteFolders")]        wizCategorySelectedTypeAllFavoriteFolders =            12,
        [helpstring("AllOEMFolders")]        wizCategorySelectedTypeAllOEMFolders =            13,
        [helpstring("OEMFolder")]        wizCategorySelectedTypeOEMFolder =            14,
        [helpstring("Groups")]        wizCategorySelectedTypeGroups =            15,
        [helpstring("Group")]        wizCategorySelectedTypeGroup =            16,
        [helpstring("GroupUnread")]        wizCategorySelectedTypeGroupUnread =            17,
        [helpstring("GroupRecent")]        wizCategorySelectedTypeGroupRecent =            18,
        [helpstring("GroupTag")]        wizCategorySelectedTypeGroupTag =            19,
        [helpstring("GroupDeleted")]        wizCategorySelectedTypeGroupDeleted =            20,
        [helpstring("Info")]        wizCategorySelectedTypeInfo =            100,
        [helpstring("Link")]        wizCategorySelectedTypeLink =            101,
    };

    enum WizBatchDownloaderJobFlags
    { 
        [helpstring("LinkTextAsTitle")]        wizWizBatchDownloaderJobLinkTextAsTitle =   0x0001,
        [helpstring("ExecuteScript")]        wizWizBatchDownloaderJobExecuteScript =        0x0002,
        [helpstring("URLAsCustomID")]        wizWizBatchDownloaderJobURLAsCustomID =        0x0004,
    };
    enum WizCommandID
    { 
        [helpstring("DocumentListCtrl Select Prev Document")]        wizCommandSelectPrevDocument =   0x0001,
        [helpstring("DocumentListCtrl Select Next Document")]        wizCommandSelectNextDocument =   0x0002,
    };

    enum WizCategoryCtrlOptions
    {
        [helpstring("ShowDocuments")]                    wizCategoryShowDocuments                =   0x0001,
        [helpstring("DisplaySubFolderDocuments")]        wizCategoryDisplaySubFolderDocuments    =   0x0002,
        [helpstring("DisplayChildTagDocuments")]        wizCategoryDisplayChildTagDocuments        =   0x0004,
        [helpstring("ShowFolders")]                        wizCategoryShowFolders                    =   0x0008,
        [helpstring("ShowTags")]                        wizCategoryShowTags                        =   0x0010,
        [helpstring("ShowStyles")]                        wizCategoryShowStyles                    =   0x0020,
        [helpstring("ShowQuickSearches")]                wizCategoryShowQuickSearches            =   0x0040,
        [helpstring("ShowFavoriteFolders")]                wizCategoryShowFavoriteFolders            =   0x0080,
        [helpstring("NoSkin")]                            wizCategoryNoSkin                        =   0x0100,
        [helpstring("NoMenu")]                            wizCategoryNoMenu                        =   0x0200,
        [helpstring("ForceBorder")]                        wizCategoryForceBorder                    =   0x0400,
        [helpstring("ShowGroups")]                        wizCategoryShowGroups                    =   0x0800,
    };
    enum WizCategoryCtrlRefreshFlags
    {
        [helpstring("RefreshAll")]                        wizCategoryRefreshAll                    =   0x0000,
        [helpstring("RefreshFolders")]                    wizCategoryRefreshFolders                =   0x0001,
        [helpstring("RefreshTags")]                        wizCategoryRefreshTags                    =   0x0002,
        [helpstring("RefreshStyles")]                    wizCategoryRefreshStyles                =   0x0004,
        [helpstring("RefreshQuickSearches")]            wizCategoryRefreshQuickSearches            =   0x0008,
        [helpstring("RefreshFavoriteFolders")]            wizCategoryRefreshFavoriteFolders        =   0x0010,
    };

    enum WizDocumentListCtrlViewStyle
    {
        [helpstring("SingleLine")]                        wizDocumentCtrlSingleLine                =   0x0001,
        [helpstring("DoubleLine")]                        wizDocumentCtrlDoubleLine                =   0x0002,
        [helpstring("MultiLine")]                        wizDocumentCtrlMultiLine                =   0x0003,
    };

    enum WizDocumentListCtrlSecondLineType
    {
        [helpstring("Auto")]                    wizDocumentCtrlSecondLineAuto                    =   0x0000,
        [helpstring("Star")]                    wizDocumentCtrlSecondLineStar                    =   0x0001,
        [helpstring("Location")]                    wizDocumentCtrlSecondLineLocation            =   0x0002,
        [helpstring("DateCreated")]                    wizDocumentCtrlSecondLineDateCreated        =   0x0003,
        [helpstring("DateModified")]                    wizDocumentCtrlSecondLineDateModified    =   0x0004,
        [helpstring("DateAccessed")]                    wizDocumentCtrlSecondLineDateAccessed    =   0x0005,
        [helpstring("URL")]                    wizDocumentCtrlSecondLineURL                        =   0x0006,
        [helpstring("Author")]                    wizDocumentCtrlSecondLineAuthor                    =   0x0007,
        [helpstring("Keywords")]                    wizDocumentCtrlSecondLineKeywords            =   0x0008,
        [helpstring("Tags")]                    wizDocumentCtrlSecondLineTags                    =   0x0009,
        [helpstring("DataSize")]                    wizDocumentCtrlSecondLineDataSize            =   0x0010,
        [helpstring("ReadCount")]                    wizDocumentCtrlSecondLineReadCount            =   0x0011,
        [helpstring("Owner")]                    wizDocumentCtrlSecondLineOwner            =   0x0012,
    };

    [
        uuid(E132C3B7-DA0C-4946-9332-D3D1822FC52C),
        helpstring("WizProgressWindow Class")
    ]
    coclass WizProgressWindow
    {
        [default] interface IWizProgressWindow;
    };
    [
        uuid(CDEF75C2-9494-4336-AF33-66980EB65E29),
        helpstring("WizSyncProgressDlg Class")
    ]
    coclass WizSyncProgressDlg
    {
        [default] interface IWizSyncProgressDlg;
    };
    [
        uuid(31B3C15F-0113-42A2-A652-93D299392ACF),
        helpstring("WizHtmlDialogExternal Class")
    ]
    coclass WizHtmlDialogExternal
    {
        [default] interface IWizHtmlDialogExternal;
    };
    [
        uuid(96D74A38-2385-41D0-A006-5D93BF796B61),
        control,
        helpstring("WizCategoryCtrl Class")
    ]
    coclass WizCategoryCtrl
    {
        [default] interface IWizCategoryCtrl;
    };
    [
        uuid(1F56B16F-6027-4F13-8277-2019548AC282),
        helpstring("WizStatusWindow Class")
    ]
    coclass WizStatusWindow
    {
        [default] interface IWizStatusWindow;
    };
};
```
