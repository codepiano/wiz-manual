```
// WizKMCore.idl : IDL source for WizKMCore
//

// This file will be processed by the MIDL tool to
// produce the type library (WizKMCore.tlb) and marshalling code.

import "oaidl.idl";
import "ocidl.idl";

[
    object,
    uuid(66EDABF2-D4D0-4B63-BFFA-EB7C53A9372A),
    dual,
    nonextensible,
    helpstring("IWizDatabase Interface"),
    pointer_default(unique)
]
interface IWizDatabase : IDispatch{
    [id(1), helpstring("method Open")] HRESULT Open([in] BSTR bstrDatabasePath);
    [id(2), helpstring("method Close")] HRESULT Close(void);
    [propget, id(3), helpstring("property DatabasePath")] HRESULT DatabasePath([out, retval] BSTR* pVal);
    [propget, id(4), helpstring("property Folders")] HRESULT Folders([out, retval] IDispatch** pVal);
    [propget, id(5), helpstring("property RootTags")] HRESULT RootTags([out, retval] IDispatch** pVal);
    [propget, id(6), helpstring("property Tags")] HRESULT Tags([out, retval] IDispatch** pVal);
    [propget, id(7), helpstring("property Styles")] HRESULT Styles([out, retval] IDispatch** pVal);
    [propget, id(8), helpstring("property Attachments")] HRESULT Attachments([out, retval] IDispatch** pVal);
    [propget, id(9), helpstring("property Metas")] HRESULT Metas([out, retval] IDispatch** pVal);
    [propget, id(10), helpstring("property MetasByName")] HRESULT MetasByName([in] BSTR bstrMetaName, [out, retval] IDispatch** pVal);
    [propget, id(11), helpstring("property Meta")] HRESULT Meta([in] BSTR bstrMetaName, [in] BSTR bstrMetaKey, [out, retval] BSTR* pVal);
    [propput, id(11), helpstring("property Meta")] HRESULT Meta([in] BSTR bstrMetaName, [in] BSTR bstrMetaKey, [in] BSTR newVal);
    [propget, id(12), helpstring("property DeletedItemsFolder")] HRESULT DeletedItemsFolder([out, retval] IDispatch** pVal);
    [id(13), helpstring("method CreateRootFolder")] HRESULT CreateRootFolder([in] BSTR bstrFolderName, [in] VARIANT_BOOL vbSync, [out,retval] IDispatch** ppNewFolderDisp);
    [id(14), helpstring("method CreateRootFolder2")] HRESULT CreateRootFolder2([in] BSTR bstrFolderName, [in] VARIANT_BOOL vbSync, [in] BSTR bstrIconFileName, [out,retval] IDispatch** ppNewFolderDisp);
    [id(15), helpstring("method CreateRootTag")] HRESULT CreateRootTag([in] BSTR bstrRootTagName, [in] BSTR bstrTagDescription, [out,retval] IDispatch** ppNewTagDisp);
    [id(16), helpstring("method CreateStyle")] HRESULT CreateStyle([in] BSTR bstrName, [in] BSTR bstrDescription, [in] LONG nTextColor, [in] LONG nBackColor, [in] VARIANT_BOOL vbTextBold, [in] LONG nFlagIndex, [out, retval] IDispatch** ppNewStyleDisp);
    [id(17), helpstring("method TagGroupFromGUID")] HRESULT TagGroupFromGUID([in] BSTR bstrTagGroupGUID, [out,retval] IDispatch** ppTagGroupDisp);
    [id(18), helpstring("method TagFromGUID")] HRESULT TagFromGUID([in] BSTR bstrTagGUID, [out,retval] IDispatch** ppTagDisp);
    [id(19), helpstring("method StyleFromGUID")] HRESULT StyleFromGUID([in] BSTR bstrStyleGUID, [out,retval] IDispatch **ppStyleDisp);
    [id(20), helpstring("method DocumentFromGUID")] HRESULT DocumentFromGUID([in] BSTR bstrDocumentGUID, [out,retval] IDispatch** ppDocumentDisp);
    [id(21), helpstring("method DocumentsFromSQL")] HRESULT DocumentsFromSQL([in] BSTR bstrSQLWhere, [out,retval] IDispatch** ppDocumentCollectionDisp);
    [id(22), helpstring("method DocumentsFromTags")] HRESULT DocumentsFromTags([in] IDispatch* pTagCollectionDisp, [out,retval] IDispatch** ppDocumentCollectionDisp);
    [id(23), helpstring("method AttachmentFromGUID")] HRESULT AttachmentFromGUID([in] BSTR bstrAttachmentGUID, [out,retval] IDispatch** ppAttachmentDisp);
    [id(24), helpstring("method GetObjectsByTime")] HRESULT GetObjectsByTime([in] DATE dtTime, [in] BSTR bstrObjectType, [out,retval] IDispatch** ppCollectionDisp);
    [id(25), helpstring("method GetModifiedObjects")] HRESULT GetModifiedObjects([in] BSTR bstrObjectType, [out,retval] IDispatch** ppCollectionDisp);
    [id(26), helpstring("method CreateTagGroupEx")] HRESULT CreateTagGroupEx([in] BSTR bstrGUID, [in] BSTR bstrName, [in] BSTR bstrDescription, [in] DATE dtModified, [in] LONGLONG nVersion, [out,retval] IDispatch** ppRetDisp);
    [id(27), helpstring("method CreateTagEx")] HRESULT CreateTagEx([in] BSTR bstrGUID, [in] BSTR bstrParentGUID, [in] BSTR bstrName, [in] BSTR bstrDescription, [in] DATE dtModified, [in] LONGLONG nVersion, [out,retval] IDispatch** ppRetDisp);
    [id(28), helpstring("method CreateStyleEx")] HRESULT CreateStyleEx([in] BSTR bstrGUID, [in] BSTR bstrName, [in] BSTR bstrDescription, [in] LONG nTextColor, [in] LONG nBackColor, [in] VARIANT_BOOL vbTextBold, [in] LONG nFlagIndex, [in] DATE dtModified, [in] LONGLONG nVersion, [out,retval] IDispatch** ppRetDisp);
    [id(29), helpstring("method CreateDocumentEx")] HRESULT CreateDocumentEx([in] BSTR bstrGUID, [in] BSTR bstrTitle, [in] BSTR bstrLocation, [in] BSTR bstrName, [in] BSTR bstrSEO, [in] BSTR bstrURL, [in] BSTR bstrAuthor, [in] BSTR bstrKeywords, [in] BSTR bstrType, [in] BSTR bstrOwner, [in] BSTR bstrFileType, [in] BSTR bstrStyleGUID, [in] DATE dtCreated, [in] DATE dtModified, [in] DATE dtAccessed, [in] LONG nIconIndex, [in] LONG nSync, [in] LONG nProtected, [in] LONG nReadCount, [in] LONG nAttachmentCount, [in] LONG nIndexed, [in] DATE dtInfoModified, [in] BSTR bstrInfoMD5, [in] DATE dtDataModified, [in] BSTR bstrDataMD5, [in] DATE dtParamModified, [in] BSTR bstrParamMD5, [in] VARIANT vTagGUIDs, [in] VARIANT vParams, [in] IUnknown* pDataStream, [in] LONGLONG nVersion, [out,retval] IDispatch** ppRetDisp);
    [id(30), helpstring("method CreateDocumentAttachmentEx")] HRESULT CreateDocumentAttachmentEx([in] BSTR bstrGUID, [in] BSTR bstrDocumentGUID, [in] BSTR bstrName, [in] BSTR bstrURL, [in] BSTR bstrDescription, [in] DATE dtInfoModified, [in] BSTR bstrInfoMD5, [in] DATE dtDataModified, [in] BSTR bstrDataMD5, [in] IUnknown* pDataStream, [in] LONGLONG nVersion, [out,retval] IDispatch** ppRetDisp);
    [id(31), helpstring("method UpdateTagGroupEx")] HRESULT UpdateTagGroupEx([in] BSTR bstrGUID, [in] BSTR bstrName, [in] BSTR bstrDescription, [in] DATE dtModified, [in] LONGLONG nVersion);
    [id(32), helpstring("method UpdateTagEx")] HRESULT UpdateTagEx([in] BSTR bstrGUID, [in] BSTR bstrParentGUID, [in] BSTR bstrName, [in] BSTR bstrDescription, [in] DATE dtModified, [in] LONGLONG nVersion);
    [id(33), helpstring("method UpdateStyleEx")] HRESULT UpdateStyleEx([in] BSTR bstrGUID, [in] BSTR bstrName, [in] BSTR bstrDescription, [in] LONG nTextColor, [in] LONG nBackColor, [in] VARIANT_BOOL vbTextBold, [in] LONG nFlagIndex, [in] DATE dtModified, [in] LONGLONG nVersion);
    [id(34), helpstring("method UpdateDocumentEx")] HRESULT UpdateDocumentEx([in] BSTR bstrGUID, [in] BSTR bstrTitle, [in] BSTR bstrLocation, [in] BSTR bstrName, [in] BSTR bstrSEO, [in] BSTR bstrURL, [in] BSTR bstrAuthor, [in] BSTR bstrKeywords, [in] BSTR bstrType, [in] BSTR bstrOwner, [in] BSTR bstrFileType, [in] BSTR bstrStyleGUID, [in] DATE dtCreated, [in] DATE dtModified, [in] DATE dtAccessed, [in] LONG nIconIndex, [in] LONG nSync, [in] LONG nProtected, [in] LONG nReadCount, [in] LONG nAttachmentCount, [in] LONG nIndexed, [in] DATE dtInfoModified, [in] BSTR bstrInfoMD5, [in] DATE dtDataModified, [in] BSTR bstrDataMD5, [in] DATE dtParamModified, [in] BSTR bstrParamMD5, [in] VARIANT vTagGUIDs, [in] VARIANT vParams, [in] IUnknown* pDataStream, [in] LONGLONG nVersion, [in] long nParts);
    [id(35), helpstring("method UpdateDocumentAttachmentEx")] HRESULT UpdateDocumentAttachmentEx([in] BSTR bstrGUID, [in] BSTR bstrDocumentGUID, [in] BSTR bstrName, [in] BSTR bstrURL, [in] BSTR bstrDescription, [in] DATE dtInfoModified, [in] BSTR bstrInfoMD5, [in] DATE dtDataModified, [in] BSTR bstrDataMD5, [in] IUnknown* pDataStream, [in] LONGLONG nVersion, [in] long nParts);
    [id(36), helpstring("method DeleteObject")] HRESULT DeleteObject([in] BSTR bstrGUID, [in] BSTR bstrType, [in] DATE dtDeleted);
    [id(37), helpstring("method ObjectExists")] HRESULT ObjectExists([in] BSTR bstrGUID, [in] BSTR bstrType, [out,retval] VARIANT_BOOL* pvbExists);
    [id(38), helpstring("method SearchDocuments")] HRESULT SearchDocuments([in] BSTR bstrKeywords, [in] LONG nFlags, [in] IDispatch* pFolderDisp, [in] LONG nMaxResults, [out,retval] IDispatch** ppDocumentCollectionDisp);
    [id(39), helpstring("method EmptyDeletedItems")] HRESULT EmptyDeletedItems();
    [id(40), helpstring("method BeginUpdate")] HRESULT BeginUpdate();
    [id(41), helpstring("method EndUpdate")] HRESULT EndUpdate();
    [id(42), helpstring("method GetRecentDocuments")] HRESULT GetRecentDocuments([in] BSTR bstrDocumentType, [in] LONG nCount, [out,retval] IDispatch** ppDocumentCollectionDisp);
    [id(43), helpstring("method GetCalendarEvents")] HRESULT GetCalendarEvents([in] DATE dtStart, [in] DATE dtEnd, [out,retval] IDispatch** ppDocumentCollectionDisp);
    [id(44), helpstring("method GetFolderByLocation")] HRESULT GetFolderByLocation([in] BSTR bstrLocation, [in] VARIANT_BOOL vbCreate, [out,retval] IDispatch** ppFolderDisp);
    [id(45), helpstring("method GetTagByName")] HRESULT GetTagByName([in] BSTR bstrTagName, [in] VARIANT_BOOL vbCreate, [in] BSTR bstrRootTagForCreate, [out,retval] IDispatch** ppTagDisp);
    [id(46), helpstring("method GetRootTagByName")] HRESULT GetRootTagByName([in] BSTR bstrTagName, [in] VARIANT_BOOL vbCreate, [out,retval] IDispatch** ppTagDisp);
    [id(47), helpstring("method GetAllDocuments")] HRESULT GetAllDocuments([out,retval] IDispatch** ppDocumentCollectionDisp);
    [id(48), helpstring("method IsModified")] HRESULT IsModified([out,retval] VARIANT_BOOL* pvbModified);
    [id(49), helpstring("method BackupIndex")] HRESULT BackupIndex();
    [id(50), helpstring("method BackupAll")] HRESULT BackupAll([in] BSTR bstrDestPath);
    [id(51), helpstring("method FileNameToDocument")] HRESULT FileNameToDocument([in] BSTR bstrFileName, [out,retval] IDispatch** ppDocumentDisp);
    [id(52), helpstring("method GetVirtualFolders")] HRESULT GetVirtualFolders([out,retval] VARIANT* pvVirtualFolderNames);
    [id(53), helpstring("method GetVirtualFolderDocuments")] HRESULT GetVirtualFolderDocuments([in] BSTR bstrVirtualFolderName, [out,retval] IDispatch** ppDocumentDisp);
    [id(54), helpstring("method GetVirtualFolderIcon")] HRESULT GetVirtualFolderIcon([in] BSTR bstrVirtualFolderName, [out,retval] BSTR* pbstrIconFileName);
    [id(55), helpstring("method GetAllFoldersDocumentCount")] HRESULT GetAllFoldersDocumentCount([in] LONG nFlags, [out,retval] VARIANT* pvDocumentCount);
    [id(56), helpstring("method DocumentsFromURL")] HRESULT DocumentsFromURL([in] BSTR bstrURL, [out,retval] IDispatch** ppDocumentCollectionDisp);
    [id(57), helpstring("method DocumentCustomGetModified")] HRESULT DocumentCustomGetModified([in] BSTR bstrDocumentLocation, [in] BSTR bstrDocumentCustomID, [out,retval] DATE* pdtCustomModified);
    [id(58), helpstring("method DocumentCustomUpdate")] HRESULT DocumentCustomUpdate([in] BSTR bstrDocumentLocation, [in] BSTR bstrTitle, [in] BSTR bstrDocumentCustomID, [in] BSTR bstrDocumentURL, [in] DATE dtCustomModified, [in] BSTR bstrHtmlFileName, [in] VARIANT_BOOL vbAllowOverwrite, [in] LONG nUpdateDocumentFlags, [out,retval] IDispatch** ppDocumentDisp);
    [id(59), helpstring("method EnableDocumentIndexing")] HRESULT EnableDocumentIndexing([in] VARIANT_BOOL vbEnable);
    [id(60), helpstring("method DumpDatabase")] HRESULT DumpDatabase([in] LONG nFlags, [out,retval] VARIANT_BOOL* pvbRet);
    [id(61), helpstring("method RebuildDatabase")] HRESULT RebuildDatabase([in] LONG nFlags, [out,retval] VARIANT_BOOL* pvbRet);
    [id(62), helpstring("method GetTodoDocument")] HRESULT GetTodoDocument([in] DATE dtDate, [out,retval] IDispatch** ppDocumentDisp);
    [id(63), helpstring("method GetCalendarEvents2")] HRESULT GetCalendarEvents2([in] DATE dtStart, [in] DATE dtEnd, [out,retval] IDispatch** ppEventCollectionDisp);
    [id(64), helpstring("method GetKnownRootTagName")] HRESULT GetKnownRootTagName([in] BSTR bstrRootTagName, [out,retval] BSTR* pbstrLocalRootTagName);
    [id(65), helpstring("method SQLQuery")] HRESULT SQLQuery([in] BSTR bstrSQL, [in] BSTR bstrOptions, [out,retval] IDispatch** ppRowsetDisp);
    [id(66), helpstring("method SQLExecute")] HRESULT SQLExecute([in] BSTR bstrSQL, [in] BSTR bstrOptions, [out,retval] LONG* pnRowsAffect);
    [id(67), helpstring("method GetAllTagsDocumentCount")] HRESULT GetAllTagsDocumentCount([in] LONG nFlags, [out,retval] VARIANT* pvDocumentCount);
    [id(68), helpstring("method AddZiwFile")] HRESULT AddZiwFile([in] LONG nFlags, [in] BSTR bstrZiwFileName, [out,retval] IDispatch** ppDocumentDisp);
    [id(69), helpstring("method Open2")] HRESULT Open2([in] BSTR bstrDatabasePath, [in] BSTR bstrPassword, [in] LONG nFlags, [out,retval] BSTR* pbstrErrorMessage);
    [id(70), helpstring("method ChangePassword")] HRESULT ChangePassword([in] BSTR bstrOldPassword, [in] BSTR bstrPassword);
    [propget, id(71), helpstring("property PasswordFlags")] HRESULT PasswordFlags([out, retval] LONG* pVal);
    [propput, id(71), helpstring("property PasswordFlags")] HRESULT PasswordFlags([in] LONG newVal);
    [propget, id(72), helpstring("property UserName")] HRESULT UserName([out, retval] BSTR* pVal);
    [propput, id(72), helpstring("property UserName")] HRESULT UserName([in] BSTR newVal);
    [id(73), helpstring("method GetEncryptedPassword")] HRESULT GetEncryptedPassword([out,retval] BSTR* pbstrPassword);
    [id(74), helpstring("method GetUserPassword")] HRESULT GetUserPassword([out,retval] BSTR* pbstrPassword);
    [id(75), helpstring("method SerCert")] HRESULT SetCert([in] BSTR bstrN, [in] BSTR bstre, [in] BSTR bstrEcnryptedd, [in] BSTR bstrHint, [in] LONG nFlags, [in] LONGLONG nWindowHandle);
    [id(76), helpstring("method GetCertN")] HRESULT GetCertN([out,retval] BSTR* pbstrN);
    [id(77), helpstring("method GetCerte")] HRESULT GetCerte([out,retval] BSTR* pbstre);
    [id(78), helpstring("method GetEncryptedCertd")] HRESULT GetEncryptedCertd([out,retval] BSTR* pbstrEncrypted_d);
    [id(79), helpstring("method GetCertHint")] HRESULT GetCertHint([out,retval] BSTR* pbstrHint);
    [id(80), helpstring("method InitCertFromServer")] HRESULT InitCertFromServer([in] LONGLONG nWindowHandle);
    [propget, id(81), helpstring("property CertPassword")] HRESULT CertPassword([out, retval] BSTR* pVal);
    [propput, id(81), helpstring("property CertPassword")] HRESULT CertPassword([in] BSTR newVal);
    [propget, id(82), helpstring("property AllDocumentsTitle")] HRESULT AllDocumentsTitle([out, retval] BSTR* pVal);
    [id(83), helpstring("method DocumentsFromTitle")] HRESULT DocumentsFromTitle([in] BSTR bstrTitle, [in] LONG nFlags, [out,retval] IDispatch** ppDocumentCollectionDisp);
    [id(84), helpstring("method DocumentsFromTagWithChildren")] HRESULT DocumentsFromTagWithChildren([in] IDispatch* pRootTagDisp, [out,retval] IDispatch** ppDocumentCollectionDisp);
    [id(85), helpstring("method GetRecentDocuments2")] HRESULT GetRecentDocuments2([in] LONG nFlags, [in] BSTR bstrDocumentType, [in] LONG nCount, [out,retval] IDispatch** ppDocumentCollectionDisp);
    [id(86), helpstring("method CreateTodo2Document")] HRESULT CreateTodo2Document([in] BSTR bstrLocation, [in] BSTR bstrTitle, [out,retval] IDispatch** ppDocumentDisp);
    [id(87), helpstring("method MoveCompletedTodoItems")] HRESULT MoveCompletedTodoItems([in] IDispatch* pSrcTodoDocumentDisp, [in] IDispatch* pDestTodoDocumentDisp);
    [id(88), helpstring("method GetTodo2Documents")] HRESULT GetTodo2Documents([in] BSTR bstrLocation, [out, retval] IDispatch** ppDocumentCollectionDisp);
    [id(89), helpstring("method DocumentsFromTagsInFolder")] HRESULT DocumentsFromTagsInFolder([in] IDispatch* pFolderDisp, [in] IDispatch* pTagsCollDisp, [out, retval] IDispatch** ppDocumentCollectionDisp);
    [id(90), helpstring("method DocumentsFromStyleInFolder")] HRESULT DocumentsFromStyleInFolder([in] IDispatch* pFolderDisp, [in] IDispatch* pStyleDisp, [out, retval] IDispatch** ppDocumentCollectionDisp);
    [id(91), helpstring("method DocumentsFromSQL2")] HRESULT DocumentsFromSQL2([in] BSTR bstrSQLWhere, [in] BSTR bstrOptions, [out,retval] IDispatch** ppDocumentCollectionDisp);
    [propget, id(92), helpstring("property AllLocations")] HRESULT AllLocations([out, retval] VARIANT* pVal);
    [propget, id(93), helpstring("property IsCustomSorted")] HRESULT IsCustomSorted([out, retval] VARIANT_BOOL* pVal);
    [id(94), helpstring("method DocumentsFromTagsInFolder2")] HRESULT DocumentsFromTagsInFolder2([in] VARIANT_BOOL vbIncludeChildTags, [in] VARIANT_BOOL vbTagsAnd, [in] IDispatch* pFolderDisp, [in] IDispatch* pTagsCollDisp, [out, retval] IDispatch** ppDocumentCollectionDisp);
    [id(95), helpstring("method RegisterLogWindow")] HRESULT RegisterLogWindow([in] LONGLONG hwndWindow);
    [id(96), helpstring("method RepairIndex")] HRESULT RepairIndex([in] BSTR bstrDestFileName);
    [id(97), helpstring("method GetMeta")] HRESULT GetMeta([in] BSTR bstrMetaName, [in] BSTR bstrMetaKey, [out,retval] BSTR* pbstrVal);
    [id(98), helpstring("method SetMeta")] HRESULT SetMeta([in] BSTR bstrMetaName, [in] BSTR bstrMetaKey, [in] BSTR newVal);
    [id(99), helpstring("method Open3")] HRESULT Open3([in] BSTR bstrDatabasePath, [in] BSTR bstrPassword, [in] BSTR bstrKbGUID, [in] LONG nFlags, [out,retval] BSTR* pbstrErrorMessage);
    [propget, id(100), helpstring("property KbGUID")] HRESULT KbGUID([out, retval] BSTR* pVal);
    [id(101), helpstring("method GetObjectsNeedToBeDownloaded")] HRESULT GetObjectsNeedToBeDownloaded([in] BSTR bstrObjectType, [in] BSTR bstrExtCondition, [out,retval] IDispatch** ppCollectionDisp);
    [propget, id(102), helpstring("property IsGroupDatabase")] HRESULT IsGroupDatabase([out, retval] VARIANT_BOOL* pVal);
    [id(103), helpstring("method SetObjectDownloaded")] HRESULT SetObjectDownloaded([in] BSTR bstrGUID, [in] BSTR bstrObjectType, [in] VARIANT_BOOL vbDownloaded);
    [id(104), helpstring("method GetCommonUsedTags")] HRESULT GetCommonUsedTags([in] LONG count, [out,retval] IDispatch** ppDisp);
    [id(105), helpstring("method GetAllTagsName")] HRESULT GetAllTagsName([out,retval] BSTR* pbstrAllTagsName);
    [propget, id(106), helpstring("property UserGroup")] HRESULT UserGroup([out, retval] LONG* pVal);
    [propget, id(107), helpstring("property UnreadCount")] HRESULT UnreadCount([out, retval] LONG* pVal);
    [id(108), helpstring("method MarkAllRead")] HRESULT MarkAllRead();
    [id(109), helpstring("method GetAllTagsDocumentCount2")] HRESULT GetAllTagsDocumentCount2([in] IDispatch* pFolderDisp, [in] LONG nFlags, [out,retval] VARIANT* pvDocumentCount);
    [propget, id(110), helpstring("property GroupName")] HRESULT GroupName([out,retval] BSTR* pVal);
    [id(111), helpstring("method CreateWizObject")] HRESULT CreateWizObject([in] BSTR bstrObjectName, [out,retval] IDispatch** ppVal);
};

[
    object,
    uuid(9B636DAC-CCF9-481E-8519-5662F35C6B0D),
    dual,
    nonextensible,
    helpstring("IWizTagCollection Interface"),
    pointer_default(unique)
]
interface IWizTagCollection : IDispatch{
    [id(DISPID_NEWENUM), propget] HRESULT _NewEnum([out, retval] IUnknown** ppUnk);
    [id(DISPID_VALUE), propget] HRESULT Item([in] long Index, [out, retval] IDispatch** pVal);
    [id(1), propget] HRESULT Count([out, retval] long * pVal);
    [id(2), helpstring("method Add")] HRESULT Add([in] IDispatch* pTagDisp);
};
[
    object,
    uuid(C710180C-D3F6-4F25-B981-B6820B01C67C),
    dual,
    nonextensible,
    helpstring("IWizTag Interface"),
    pointer_default(unique)
]
interface IWizTag : IDispatch{
    [propget, id(1), helpstring("property GUID")] HRESULT GUID([out, retval] BSTR* pVal);
    [propget, id(2), helpstring("property ParentTag")] HRESULT ParentTag([out, retval] IDispatch** pVal);
    [propput, id(2), helpstring("property ParentTag")] HRESULT ParentTag([in ] IDispatch* newVal);
    [propget, id(3), helpstring("property Name")] HRESULT Name([out, retval] BSTR* pVal);
    [propput, id(3), helpstring("property Name")] HRESULT Name([in] BSTR newVal);
    [propget, id(4), helpstring("property Description")] HRESULT Description([out, retval] BSTR* pVal);
    [propput, id(4), helpstring("property Description")] HRESULT Description([in] BSTR newVal);
    [propget, id(5), helpstring("property DateModified")] HRESULT DateModified([out, retval] DATE* pVal);
    [propget, id(6), helpstring("property Documents")] HRESULT Documents([out, retval] IDispatch** pVal);
    [propget, id(7), helpstring("property ParentTagGUID")] HRESULT ParentTagGUID([out, retval] BSTR* pVal);
    [propget, id(8), helpstring("property Version")] HRESULT Version([out, retval] LONGLONG* pVal);
    [propput, id(8), helpstring("property Version")] HRESULT Version([in] LONGLONG newVal);
    [propget, id(9), helpstring("property Database")] HRESULT Database([out, retval] IDispatch** pVal);
    [id(10), helpstring("method Delete")] HRESULT Delete(void);
    [id(11), helpstring("method Reload")] HRESULT Reload(void);
    [propget, id(12), helpstring("property Children")] HRESULT Children([out, retval] IDispatch** pVal);
    [id(13), helpstring("method CreateChildTag")] HRESULT CreateChildTag([in] BSTR bstrTagName, [in] BSTR bstrTagDescription, [out,retval] IDispatch** ppNewTagDisp);
    [id(14), helpstring("method MoveToRoot")] HRESULT MoveToRoot();
    [id(15), helpstring("method IsIn")] HRESULT IsIn([in] IDispatch* pFolderDisp, [out,retval] VARIANT_BOOL* pvbRet);
    [propget, id(16), helpstring("property DisplayName")] HRESULT DisplayName([out, retval] BSTR* pVal);
    [propget, id(17), helpstring("property ParentGUID")] HRESULT ParentGUID([out, retval] BSTR* pVal);
    [propget, id(18), helpstring("property FullPath")] HRESULT FullPath([out, retval] BSTR* pVal);
    [id(19), helpstring("method GetChildTagByPath")] HRESULT GetChildTagByPath([in] BSTR bstrTagPath, [in] VARIANT_BOOL vbCreate, [out,retval] IDispatch** ppNewTagDisp);
    [id(20), helpstring("method GetPinYin")] HRESULT GetPinYin([in] LONG flags, [out,retval] BSTR* pbstrRetPinYin);
    [propget, id(21), helpstring("property AllChildren")] HRESULT AllChildren([out, retval] IDispatch** pVal);
};
[
    object,
    uuid(F8AE7F18-3C30-47D1-8902-947E5676895C),
    dual,
    nonextensible,
    helpstring("IWizDocumentCollection Interface"),
    pointer_default(unique)
]
interface IWizDocumentCollection : IDispatch{
    [id(DISPID_NEWENUM), propget] HRESULT _NewEnum([out, retval] IUnknown** ppUnk);
    [id(DISPID_VALUE), propget] HRESULT Item([in] long Index, [out, retval] IDispatch** pVal);
    [id(1), propget] HRESULT Count([out, retval] long * pVal);
    [id(2), helpstring("method Add")] HRESULT Add([in] IDispatch* pDocumentDisp);
    [id(3), helpstring("method Delete")] HRESULT Delete([in] LONG nIndex);
    [id(4), helpstring("method Clear")] HRESULT Clear();    
};
[
    object,
    uuid(5B78A65E-6EE8-41C7-B1DA-8ECBF9D917B0),
    dual,
    nonextensible,
    helpstring("IWizDocument Interface"),
    pointer_default(unique)
]
interface IWizDocument : IDispatch{
    [propget, id(1), helpstring("property GUID")] HRESULT GUID([out, retval] BSTR* pVal);
    [propget, id(2), helpstring("property Title")] HRESULT Title([out, retval] BSTR* pVal);
    [propput, id(2), helpstring("property Title")] HRESULT Title([in] BSTR newVal);
    [propget, id(3), helpstring("property Author")] HRESULT Author([out, retval] BSTR* pVal);
    [propput, id(3), helpstring("property Author")] HRESULT Author([in] BSTR newVal);
    [propget, id(4), helpstring("property Keywords")] HRESULT Keywords([out, retval] BSTR* pVal);
    [propput, id(4), helpstring("property Keywords")] HRESULT Keywords([in] BSTR newVal);
    [propget, id(5), helpstring("property Tags")] HRESULT Tags([out, retval] IDispatch** pVal);
    [propput, id(5), helpstring("property Tags")] HRESULT Tags([in] IDispatch* newVal);
    [propget, id(6), helpstring("property Name")] HRESULT Name([out, retval] BSTR* pVal);
    [propput, id(6), helpstring("property Name")] HRESULT Name([in] BSTR newVal);
    [propget, id(7), helpstring("property Location")] HRESULT Location([out, retval] BSTR* pVal);
    [propget, id(8), helpstring("property FileName")] HRESULT FileName([out, retval] BSTR* pVal);
    [propget, id(9), helpstring("property SEO")] HRESULT SEO([out, retval] BSTR* pVal);
    [propput, id(9), helpstring("property SEO")] HRESULT SEO([in] BSTR newVal);
    [propget, id(10), helpstring("property URL")] HRESULT URL([out, retval] BSTR* pVal);
    [propput, id(10), helpstring("property URL")] HRESULT URL([in] BSTR newVal);
    [propget, id(11), helpstring("property Type")] HRESULT Type([out, retval] BSTR* pVal);
    [propput, id(11), helpstring("property Type")] HRESULT Type([in] BSTR newVal);
    [propget, id(12), helpstring("property Owner")] HRESULT Owner([out, retval] BSTR* pVal);
    [propput, id(12), helpstring("property Owner")] HRESULT Owner([in] BSTR newVal);
    [propget, id(13), helpstring("property FileType")] HRESULT FileType([out, retval] BSTR* pVal);
    [propput, id(13), helpstring("property FileType")] HRESULT FileType([in] BSTR newVal);
    [propget, id(14), helpstring("property Style")] HRESULT Style([out, retval] IDispatch** pVal);
    [propput, id(14), helpstring("property Style")] HRESULT Style([in] IDispatch* newVal);
    [propget, id(15), helpstring("property IconIndex")] HRESULT IconIndex([out, retval] LONG* pVal);
    [propput, id(15), helpstring("property IconIndex")] HRESULT IconIndex([in] LONG newVal);
    [propget, id(16), helpstring("property Sync")] HRESULT Sync([out, retval] VARIANT_BOOL* pVal);
    [propput, id(16), helpstring("property Sync")] HRESULT Sync([in] VARIANT_BOOL newVal);
    [propget, id(17), helpstring("property Protect")] HRESULT Protect([out, retval] LONG* pVal);
    [propput, id(17), helpstring("property Protect")] HRESULT Protect([in] LONG newVal);
    [propget, id(18), helpstring("property ReadCount")] HRESULT ReadCount([out, retval] LONG* pVal);
    [propput, id(18), helpstring("property ReadCount")] HRESULT ReadCount([in] LONG newVal);
    [propget, id(19), helpstring("property AttachmentCount")] HRESULT AttachmentCount([out, retval] LONG* pVal);
    [propget, id(20), helpstring("property Indexed")] HRESULT Indexed([out, retval] VARIANT_BOOL* pVal);
    [propput, id(20), helpstring("property Indexed")] HRESULT Indexed([in] VARIANT_BOOL newVal);
    [propget, id(21), helpstring("property DateCreated")] HRESULT DateCreated([out, retval] DATE* pVal);
    [propput, id(21), helpstring("property DateCreated")] HRESULT DateCreated([in] DATE newVal);
    [propget, id(22), helpstring("property DateModified")] HRESULT DateModified([out, retval] DATE* pVal);
    [propput, id(22), helpstring("property DateModified")] HRESULT DateModified([in] DATE newVal);
    [propget, id(23), helpstring("property DateAccessed")] HRESULT DateAccessed([out, retval] DATE* pVal);
    [propput, id(23), helpstring("property DateAccessed")] HRESULT DateAccessed([in] DATE newVal);
    [propget, id(24), helpstring("property InfoDateModified")] HRESULT InfoDateModified([out, retval] DATE* pVal);
    [propget, id(25), helpstring("property InfoMD5")] HRESULT InfoMD5([out, retval] BSTR* pVal);
    [propget, id(26), helpstring("property DataDateModified")] HRESULT DataDateModified([out, retval] DATE* pVal);
    [propget, id(27), helpstring("property DataMD5")] HRESULT DataMD5([out, retval] BSTR* pVal);
    [propget, id(28), helpstring("property ParamDateModified")] HRESULT ParamDateModified([out, retval] DATE* pVal);
    [propget, id(29), helpstring("property ParamMD5")] HRESULT ParamMD5([out, retval] BSTR* pVal);
    [propget, id(30), helpstring("property Attachments")] HRESULT Attachments([out, retval] IDispatch** pVal);
    [propget, id(31), helpstring("property ParamValue")] HRESULT ParamValue([in] BSTR bstrParamName, [out, retval] BSTR* pVal);
    [propput, id(31), helpstring("property ParamValue")] HRESULT ParamValue([in] BSTR bstrParamName, [in] BSTR newVal);
    [propget, id(32), helpstring("property Params")] HRESULT Params([out, retval] IDispatch** pVal);
    [propget, id(33), helpstring("property Parent")] HRESULT Parent([out, retval] IDispatch** pVal);
    [propget, id(34), helpstring("property StyleGUID")] HRESULT StyleGUID([out, retval] BSTR* pVal);
    [propput, id(34), helpstring("property StyleGUID")] HRESULT StyleGUID([in] BSTR newVal);
    [propget, id(35), helpstring("property Version")] HRESULT Version([out, retval] LONGLONG* pVal);
    [propput, id(35), helpstring("property Version")] HRESULT Version([in] LONGLONG newVal);
    [propget, id(36), helpstring("property TagsText")] HRESULT TagsText([out, retval] BSTR* pVal);
    [propput, id(36), helpstring("property TagsText")] HRESULT TagsText([in] BSTR newVal);
    [propget, id(37), helpstring("property Database")] HRESULT Database([out, retval] IDispatch** pVal);
    [propget, id(38), helpstring("property FilePath")] HRESULT FilePath([out, retval] BSTR* pVal);
    [propget, id(39), helpstring("property AttachmentsFilePath")] HRESULT AttachmentsFilePath([out, retval] BSTR* pVal);
    [id(40), helpstring("method AddTag")] HRESULT AddTag([in] IDispatch* pTagDisp);
    [id(41), helpstring("method RemoveTag")] HRESULT RemoveTag([in] IDispatch* pTagDisp);
    [id(42), helpstring("method RemoveAllTags")] HRESULT RemoveAllTags();
    [id(43), helpstring("method Delete")] HRESULT Delete(void);
    [id(44), helpstring("method Reload")] HRESULT Reload(void);
    [id(45), helpstring("method AddAttachment")] HRESULT AddAttachment([in] BSTR bstrFileName, [out,retval] IDispatch** ppNewAttachmentDisp);
    [id(46), helpstring("method SaveToHtml")] HRESULT SaveToHtml([in] BSTR bstrHtmlFileName, [in] LONG nFlags);
    [id(47), helpstring("method MoveTo")] HRESULT MoveTo([in] IDispatch* pDestFolderDisp);
    [id(48), helpstring("method CopyTo")] HRESULT CopyTo([in] IDispatch* pDestFolderDisp, [out,retval] IDispatch** ppNewDocumentDisp);
    [id(49), helpstring("method IsIn")] HRESULT IsIn([in] IDispatch* pFolderDisp, [out,retval] VARIANT_BOOL* pvbRet);
    [id(50), helpstring("method UpdateDocument")] HRESULT UpdateDocument([in] BSTR bstrHtmlFileName, [in] LONG nFlags);
    [id(51), helpstring("method UpdateDocument2")] HRESULT UpdateDocument2([in] IDispatch* pIHTMLDocument2Disp, [in] LONG nFlags);
    [id(52), helpstring("method UpdateDocument3")] HRESULT UpdateDocument3([in] BSTR bstrHtml, [in] LONG nFlags);
    [id(53), helpstring("method UpdateDocument4")] HRESULT UpdateDocument4([in] BSTR bstrHtml, [in] BSTR bstrURL, [in] LONG nFlags);
    [id(54), helpstring("method UpdateDocument5")] HRESULT UpdateDocument5([in] BSTR bstrHtmlFileName);
    [id(55), helpstring("method UpdateDocument6")] HRESULT UpdateDocument6([in] BSTR bstrHtmlFileName, [in] BSTR bstrURL, [in] LONG nFlags);
    [id(56), helpstring("method CompareTo")] HRESULT CompareTo([in] IDispatch* pDocumentDisp, [in] LONG nCompareBy, [out,retval] LONG* pnRet);
    [id(57), helpstring("method GetText")] HRESULT GetText([in] UINT nFlags, [out, retval] BSTR* pbstrText);
    [id(58), helpstring("method GetHtml")] HRESULT GetHtml([out, retval] BSTR* pbstrHtml);
    [id(59), helpstring("method BeginUpdateParams")] HRESULT BeginUpdateParams();
    [id(60), helpstring("method EndUpdateParams")] HRESULT EndUpdateParams();
    [id(61), helpstring("method AddToCalendar")] HRESULT AddToCalendar([in] DATE dtStart, [in] DATE dtEnd, [in] BSTR bstrExtInfo);
    [id(62), helpstring("method RemoveFromCalendar")] HRESULT RemoveFromCalendar(void);
    [id(63), helpstring("method ChangeTitleAndFileName")] HRESULT ChangeTitleAndFileName([in] BSTR bstrTitle);
    [id(64), helpstring("method GetIconFileName")] HRESULT GetIconFileName([out,retval] BSTR* pbstrIconFileName);
    [propget, id(65), helpstring("property TodoItems")] HRESULT TodoItems([out, retval] IDispatch** pVal);
    [propput, id(65), helpstring("property TodoItems")] HRESULT TodoItems([in] IDispatch* newVal);
    [propget, id(66), helpstring("property Event")] HRESULT Event([out, retval] IDispatch** pVal);
    [propput, id(66), helpstring("property Event")] HRESULT Event([in] IDispatch* newVal);
    [id(67), helpstring("method AddToCalendar2")] HRESULT AddToCalendar2([in] DATE dtStart, [in] DATE dtEnd, [in] BSTR bstrRecurrence, [in] BSTR bstrEndRecurrence, [in] BSTR bstrExtInfo);
    [id(68), helpstring("method SetTagsText2")] HRESULT SetTagsText2([in] BSTR bstrTagsText, BSTR bstrParentTagName);
    [id(69), helpstring("method PermanentlyDelete")] HRESULT PermanentlyDelete(void);
    [id(70), helpstring("method GetProtectedByData")] HRESULT GetProtectedByData([out,retval] LONG* pnProtected);
    [id(71), helpstring("method AutoEncrypt")] HRESULT AutoEncrypt([in] VARIANT_BOOL vbDecrypt);
    [propget, id(72), helpstring("property Flags")] HRESULT Flags([out, retval] LONG* pVal);
    [propput, id(72), helpstring("property Flags")] HRESULT Flags([in] LONG newVal);
    [propget, id(73), helpstring("property AlwaysOnTop")] HRESULT AlwaysOnTop([out, retval] VARIANT_BOOL* pVal);
    [propput, id(73), helpstring("property AlwaysOnTop")] HRESULT AlwaysOnTop([in] VARIANT_BOOL newVal);
    [propget, id(74), helpstring("property Rate")] HRESULT Rate([out, retval] LONG* pVal);
    [propput, id(74), helpstring("property Rate")] HRESULT Rate([in] LONG newVal);
    //[propget, id(75), helpstring("property SystemTags")] HRESULT SystemTags([out, retval] BSTR* pVal);
    //[propput, id(75), helpstring("property SystemTags")] HRESULT SystemTags([in] BSTR newVal);
    //[id(76), helpstring("method IsShared")] HRESULT IsShared([out,retval] VARIANT_BOOL* pvbShared);
    //[id(77), helpstring("method SetShared")] HRESULT SetShared([in] VARIANT_BOOL vbShared);
    [id(78), helpstring("method SetDataDateModified")] HRESULT SetDataDateModified([in] DATE dtDataModified);
    //[propget, id(79), helpstring("property ShareFlags")] HRESULT ShareFlags([out, retval] LONG* pVal);
    [id(80), helpstring("method GetParamValue")] HRESULT GetParamValue([in] BSTR bstrParamName, [out,retval] BSTR* pbstrVal);
    [id(81), helpstring("method SetParamValue")] HRESULT SetParamValue([in] BSTR bstrParamName, [in] BSTR newVal);
    [propget, id(82), helpstring("property Downloaded")] HRESULT Downloaded([out, retval] VARIANT_BOOL* pVal);
    [propput, id(82), helpstring("property Downloaded")] HRESULT Downloaded([in] VARIANT_BOOL newVal);
    [id(83), helpstring("method CheckDocumentData")] HRESULT CheckDocumentData([in] LONGLONG hwndParent, [out,retval] VARIANT_BOOL* vbRet);
    [id(84), helpstring("method SetData")] HRESULT SetData([in] IUnknown* pData);
    [id(85), helpstring("method GetData")] HRESULT GetData([in] IUnknown** ppData);
    [propget, id(86), helpstring("property AbstractText")] HRESULT AbstractText([out, retval] BSTR* pVal);
    [propget, id(87), helpstring("property AbstractImage")] HRESULT AbstractImage([out, retval] LONGLONG* pVal);
    [propget, id(88), helpstring("property AbstractImageOwner")] HRESULT AbstractImageOwner([out, retval] LONGLONG* pVal);
    [propget, id(89), helpstring("property CanEdit")] HRESULT CanEdit([out, retval] VARIANT_BOOL* pVal);
    [propget, id(90), helpstring("property ServerVersion")] HRESULT ServerVersion([out, retval] LONGLONG* pVal);
    [propput, id(90), helpstring("property ServerVersion")] HRESULT ServerVersion([in] LONGLONG newVal);
    [id(91), helpstring("method GetLocalFlags")] HRESULT GetLocalFlags([out,retval] LONG* flags);
    [id(92), helpstring("method SetLocalFlags")] HRESULT SetLocalFlags([in] LONG flags);
    [id(93), helpstring("method SetServerDataInfo")] HRESULT SetServerDataInfo([in] DATE tDataModified, [in] BSTR bstrDataMD5);
};
[
    object,
    uuid(9B1E33B1-C799-4F7F-8926-54A512D45635),
    dual,
    nonextensible,
    helpstring("IWizDocumentAttachmentCollection Interface"),
    pointer_default(unique)
]
interface IWizDocumentAttachmentCollection : IDispatch{
    [id(DISPID_NEWENUM), propget] HRESULT _NewEnum([out, retval] IUnknown** ppUnk);
    [id(DISPID_VALUE), propget] HRESULT Item([in] long Index, [out, retval] IDispatch** pVal);
    [id(1), propget] HRESULT Count([out, retval] long * pVal);
    [id(2), helpstring("method Add")] HRESULT Add([in] IDispatch* pDocumentAttachmentDisp);
};
[
    object,
    uuid(9368D723-8B76-45C3-9B65-BD88B0A1BE64),
    dual,
    nonextensible,
    helpstring("IWizDocumentAttachment Interface"),
    pointer_default(unique)
]
interface IWizDocumentAttachment : IDispatch{
    [propget, id(1), helpstring("property GUID")] HRESULT GUID([out, retval] BSTR* pVal);
    [propget, id(2), helpstring("property Name")] HRESULT Name([out, retval] BSTR* pVal);
    [propput, id(2), helpstring("property Name")] HRESULT Name([in] BSTR newVal);
    [propget, id(3), helpstring("property Description")] HRESULT Description([out, retval] BSTR* pVal);
    [propput, id(3), helpstring("property Description")] HRESULT Description([in] BSTR newVal);
    [propget, id(4), helpstring("property URL")] HRESULT URL([out, retval] BSTR* pVal);
    [propput, id(4), helpstring("property URL")] HRESULT URL([in] BSTR newVal);
    [propget, id(5), helpstring("property Size")] HRESULT Size([out, retval] LONG* pVal);
    [propget, id(6), helpstring("property InfoDateModified")] HRESULT InfoDateModified([out, retval] DATE* pVal);
    [propget, id(7), helpstring("property InfoMD5")] HRESULT InfoMD5([out, retval] BSTR* pVal);
    [propget, id(8), helpstring("property DataDateModified")] HRESULT DataDateModified([out, retval] DATE* pVal);
    [propget, id(9), helpstring("property DataMD5")] HRESULT DataMD5([out, retval] BSTR* pVal);
    [propget, id(10), helpstring("property Document")] HRESULT Document([out, retval] IDispatch** pVal);
    [propget, id(11), helpstring("property DocumentGUID")] HRESULT DocumentGUID([out, retval] BSTR* pVal);
    [propget, id(12), helpstring("property FileName")] HRESULT FileName([out, retval] BSTR* pVal);
    [propget, id(13), helpstring("property Version")] HRESULT Version([out, retval] LONGLONG* pVal);
    [propput, id(13), helpstring("property Version")] HRESULT Version([in] LONGLONG newVal);
    [id(14), helpstring("method Delete")] HRESULT Delete(void);
    [id(15), helpstring("method Reload")] HRESULT Reload(void);
    [id(16), helpstring("method UpdateDataMD5")] HRESULT UpdateDataMD5(void);
    [propget, id(17), helpstring("property Downloaded")] HRESULT Downloaded([out, retval] VARIANT_BOOL* pVal);
    [propput, id(17), helpstring("property Downloaded")] HRESULT Downloaded([in] VARIANT_BOOL newVal);
    [id(18), helpstring("method CheckAttachmentData")] HRESULT CheckAttachmentData([in] LONGLONG hwndParent, [out,retval] VARIANT_BOOL* vbRet);
    [id(19), helpstring("method SetData")] HRESULT SetData([in] IUnknown* pData);
    [id(20), helpstring("method GetData")] HRESULT GetData([in] IUnknown* pData);
    [propget, id(21), helpstring("property ServerVersion")] HRESULT ServerVersion([out, retval] LONGLONG* pVal);
    [propput, id(21), helpstring("property ServerVersion")] HRESULT ServerVersion([in] LONGLONG newVal);
    [id(22), helpstring("method GetLocalFlags")] HRESULT GetLocalFlags([out,retval] LONG* flags);
    [id(23), helpstring("method SetLocalFlags")] HRESULT SetLocalFlags([in] LONG flags);
    [id(93), helpstring("method SetServerDataInfo")] HRESULT SetServerDataInfo([in] DATE tDataModified, [in] BSTR bstrDataMD5);
};
[
    object,
    uuid(6A7B925B-14C7-434B-80E9-2165CD059A79),
    dual,
    nonextensible,
    helpstring("IWizFolderCollection Interface"),
    pointer_default(unique)
]
interface IWizFolderCollection : IDispatch{
    [id(DISPID_NEWENUM), propget] HRESULT _NewEnum([out, retval] IUnknown** ppUnk);
    [id(DISPID_VALUE), propget] HRESULT Item([in] long Index, [out, retval] IDispatch** pVal);
    [id(1), propget] HRESULT Count([out, retval] long * pVal);
};
[
    object,
    uuid(B0E49F7C-CF4B-4DD4-A3F3-432F628C3A31),
    dual,
    nonextensible,
    helpstring("IWizFolder Interface"),
    pointer_default(unique)
]
interface IWizFolder : IDispatch{
    [propget, id(1), helpstring("property Name")] HRESULT Name([out, retval] BSTR* pVal);
    [propput, id(1), helpstring("property Name")] HRESULT Name([in] BSTR newVal);
    [propget, id(2), helpstring("property Sync")] HRESULT Sync([out, retval] VARIANT_BOOL* pVal);
    [propput, id(2), helpstring("property Sync")] HRESULT Sync([in] VARIANT_BOOL newVal);
    [propget, id(3), helpstring("property FullPath")] HRESULT FullPath([out, retval] BSTR* pVal);
    [propget, id(4), helpstring("property Location")] HRESULT Location([out, retval] BSTR* pVal);
    [propget, id(5), helpstring("property Documents")] HRESULT Documents([out, retval] IDispatch** pVal);
    [propget, id(6), helpstring("property Folders")] HRESULT Folders([out, retval] IDispatch** pVal);
    [propget, id(7), helpstring("property IsRoot")] HRESULT IsRoot([out, retval] VARIANT_BOOL* pVal);
    [propget, id(8), helpstring("property IsDeletedItems")] HRESULT IsDeletedItems([out, retval] VARIANT_BOOL* pVal);
    [propget, id(9), helpstring("property Parent")] HRESULT Parent([out, retval] IDispatch** pVal);
    [propget, id(10), helpstring("property Database")] HRESULT Database([out, retval] IDispatch** pVal);
    [id(11), helpstring("method CreateSubFolder")] HRESULT CreateSubFolder([in] BSTR bstrFolderName, [out,retval] IDispatch** ppNewFolderDisp);
    [id(12), helpstring("method CreateDocument")] HRESULT CreateDocument([in] BSTR bstrTitle, [in] BSTR strName, [in] BSTR bstrURL, [out,retval] IDispatch** ppNewDocumentDisp);
    [id(13), helpstring("method CreateDocument2")] HRESULT CreateDocument2([in] BSTR bstrTitle, [in] BSTR bstrURL, [out,retval] IDispatch** ppNewDocumentDisp);
    [id(14), helpstring("method Delete")] HRESULT Delete(void);
    [id(15), helpstring("method MoveTo")] HRESULT MoveTo([in] IDispatch* pDestFolderDisp);
    [id(16), helpstring("method MoveToRoot")] HRESULT MoveToRoot();
    [id(17), helpstring("method IsIn")] HRESULT IsIn([in] IDispatch* pFolderDisp, [out,retval] VARIANT_BOOL* pvbRet);
    [id(18), helpstring("method SetIcon")] HRESULT SetIcon([in] LONGLONG hIcon);
    [id(19), helpstring("method SetIcon2")] HRESULT SetIcon2([in] BSTR bstrIconFileName, [in] LONG nIconIndex);
    [id(20), helpstring("method GetIconFileName")] HRESULT GetIconFileName([out,retval] BSTR* pbstrIconFileName);
    [id(21), helpstring("method GetDocumentCount")] HRESULT GetDocumentCount([in] LONG nFlags, [out,retval] LONG* pnCount);
    [id(22), helpstring("method GetDisplayName")] HRESULT GetDisplayName([in] LONG nFlags, [out,retval] BSTR* pbstrDisplayName);
    [id(23), helpstring("method GetDisplayTemplate")] HRESULT GetDisplayTemplate([out,retval] BSTR* pbstrReaingTemplateFileName);
    [id(24), helpstring("method SetDisplayTemplate")] HRESULT SetDisplayTemplate([in] BSTR bstrReaingTemplateFileName);
    [propget, id(25), helpstring("property RootFolder")] HRESULT RootFolder([out, retval] IDispatch** pVal);
    [propget, id(26), helpstring("property Settings")] HRESULT Settings([in] BSTR bstrSection, [in] BSTR bstrKeyName, [out, retval] BSTR* pVal);
    [propput, id(26), helpstring("property Settings")] HRESULT Settings([in] BSTR bstrSection, [in] BSTR bstrKeyName, [in] BSTR newVal);
    [propget, id(27), helpstring("property Encrypt")] HRESULT Encrypt([out, retval] VARIANT_BOOL* pVal);
    [propput, id(27), helpstring("property Encrypt")] HRESULT Encrypt([in] VARIANT_BOOL newVal);
    [propget, id(28), helpstring("property IsEmpty")] HRESULT IsEmpty([out, retval] VARIANT_BOOL* pVal);
    [propget, id(29), helpstring("property SortPos")] HRESULT SortPos([out, retval] LONGLONG* pVal);
    [propput, id(29), helpstring("property SortPos")] HRESULT SortPos([in] LONGLONG newVal);
    [propget, id(30), helpstring("property IsCustomSorted")] HRESULT IsCustomSorted([out, retval] VARIANT_BOOL* pVal);
    [id(31), helpstring("method CanMove")] HRESULT CanMove([in] IDispatch* pDestFolderDisp, [out, retval] VARIANT_BOOL* pVal);
    [propget, id(32), helpstring("property UsedTags")] HRESULT UsedTags([out, retval] IDispatch** pVal);
    [id(33), helpstring("method GetLocationDisplayName")] HRESULT GetLocationDisplayName([in] LONG nFlags, [out,retval] BSTR* pbstrDisplayName);
};
[
    object,
    uuid(77C6F00C-6A41-4A14-B3E6-986401D5A5D5),
    dual,
    nonextensible,
    helpstring("IWizStyle Interface"),
    pointer_default(unique)
]
interface IWizStyle : IDispatch{
    [propget, id(1), helpstring("property GUID")] HRESULT GUID([out, retval] BSTR* pVal);
    [propget, id(2), helpstring("property Name")] HRESULT Name([out, retval] BSTR* pVal);
    [propput, id(2), helpstring("property Name")] HRESULT Name([in] BSTR newVal);
    [propget, id(3), helpstring("property Description")] HRESULT Description([out, retval] BSTR* pVal);
    [propput, id(3), helpstring("property Description")] HRESULT Description([in] BSTR newVal);
    [propget, id(4), helpstring("property TextColor")] HRESULT TextColor([out, retval] LONG* pVal);
    [propput, id(4), helpstring("property TextColor")] HRESULT TextColor([in] LONG newVal);
    [propget, id(5), helpstring("property BackColor")] HRESULT BackColor([out, retval] LONG* pVal);
    [propput, id(5), helpstring("property BackColor")] HRESULT BackColor([in] LONG newVal);
    [propget, id(6), helpstring("property TextBold")] HRESULT TextBold([out, retval] VARIANT_BOOL* pVal);
    [propput, id(6), helpstring("property TextBold")] HRESULT TextBold([in] VARIANT_BOOL newVal);
    [propget, id(7), helpstring("property FlagIndex")] HRESULT FlagIndex([out, retval] LONG* pVal);
    [propput, id(7), helpstring("property FlagIndex")] HRESULT FlagIndex([in] LONG newVal);
    [propget, id(8), helpstring("property DateModified")] HRESULT DateModified([out, retval] DATE* pVal);
    [propget, id(9), helpstring("property Documents")] HRESULT Documents([out, retval] IDispatch** pVal);
    [propget, id(10), helpstring("property Version")] HRESULT Version([out, retval] LONGLONG* pVal);
    [propput, id(10), helpstring("property Version")] HRESULT Version([in] LONGLONG newVal);
    [propget, id(11), helpstring("property Database")] HRESULT Database([out, retval] IDispatch** pVal);
    [id(12), helpstring("method Delete")] HRESULT Delete(void);
    [id(13), helpstring("method Reload")] HRESULT Reload(void);
};
[
    object,
    uuid(E6A5A90D-5766-428F-BE25-AFC9D641CB88),
    dual,
    nonextensible,
    helpstring("IWizStyleCollection Interface"),
    pointer_default(unique)
]
interface IWizStyleCollection : IDispatch{
    [id(DISPID_NEWENUM), propget] HRESULT _NewEnum([out, retval] IUnknown** ppUnk);
    [id(DISPID_VALUE), propget] HRESULT Item([in] long Index, [out, retval] IDispatch** pVal);
    [id(1), propget] HRESULT Count([out, retval] long * pVal);
};
[
    object,
    uuid(19F99E93-2C32-411E-A407-2349716D0221),
    dual,
    nonextensible,
    helpstring("IWizKMProtocol Interface"),
    pointer_default(unique)
]
interface IWizKMProtocol : IDispatch{
};
[
    object,
    uuid(46C5786E-A72E-4089-B4DB-70B3DD174FEB),
    dual,
    nonextensible,
    helpstring("IWizDocumentParam Interface"),
    pointer_default(unique)
]
interface IWizDocumentParam : IDispatch{
    [propget, id(1), helpstring("property Name")] HRESULT Name([out, retval] BSTR* pVal);
    [propget, id(2), helpstring("property Value")] HRESULT Value([out, retval] BSTR* pVal);
    [propget, id(3), helpstring("property Document")] HRESULT Document([out, retval] IDispatch** pVal);
    [id(4), helpstring("method Delete")] HRESULT Delete(void);
};
[
    object,
    uuid(8DB6192A-1C66-4DD0-ABD4-E80C12969A13),
    dual,
    nonextensible,
    helpstring("IWizDocumentParamCollection Interface"),
    pointer_default(unique)
]
interface IWizDocumentParamCollection : IDispatch{
    [id(DISPID_NEWENUM), propget] HRESULT _NewEnum([out, retval] IUnknown** ppUnk);
    [id(DISPID_VALUE), propget] HRESULT Item([in] long Index, [out, retval] IDispatch** pVal);
    [id(1), propget] HRESULT Count([out, retval] long * pVal);
};
[
    object,
    uuid(8FEE8441-A050-4CF7-99B6-4821DB945D80),
    dual,
    nonextensible,
    helpstring("IWizMeta Interface"),
    pointer_default(unique)
]
interface IWizMeta : IDispatch{
    [propget, id(1), helpstring("property Name")] HRESULT Name([out, retval] BSTR* pVal);
    [propget, id(2), helpstring("property Key")] HRESULT Key([out, retval] BSTR* pVal);
    [propget, id(3), helpstring("property Value")] HRESULT Value([out, retval] BSTR* pVal);
    [propput, id(3), helpstring("property Value")] HRESULT Value([in] BSTR newVal);
    [propget, id(4), helpstring("property DateModified")] HRESULT DateModified([out, retval] DATE* pVal);
    [id(5), helpstring("method Delete")] HRESULT Delete(void);
};
[
    object,
    uuid(7492D632-5461-42BC-BF59-6D93393892F8),
    dual,
    nonextensible,
    helpstring("IWizMetaCollection Interface"),
    pointer_default(unique)
]
interface IWizMetaCollection : IDispatch{
    [id(DISPID_NEWENUM), propget] HRESULT _NewEnum([out, retval] IUnknown** ppUnk);
    [id(DISPID_VALUE), propget] HRESULT Item([in] long Index, [out, retval] IDispatch** pVal);
    [id(1), propget] HRESULT Count([out, retval] long * pVal);
};
[
    object,
    uuid(E3C896AD-3E1F-453E-800C-07A1837E2F23),
    dual,
    nonextensible,
    helpstring("IWizSettings Interface"),
    pointer_default(unique)
]
interface IWizSettings : IDispatch{
    [id(1), helpstring("method Open")] HRESULT Open([in] BSTR bstrFileName);
    [id(2), helpstring("method Close")] HRESULT Close(void);
    [propget, id(3), helpstring("property IsDirty")] HRESULT IsDirty([out, retval] VARIANT_BOOL* pVal);
    [propget, id(4), helpstring("property StringValue")] HRESULT StringValue([in] BSTR bstrSection, [in] BSTR bstrKey, [out, retval] BSTR* pVal);
    [propput, id(4), helpstring("property StringValue")] HRESULT StringValue([in] BSTR bstrSection, [in] BSTR bstrKey, [in] BSTR newVal);
    [propget, id(5), helpstring("property IntValue")] HRESULT IntValue([in] BSTR bstrSection, [in] BSTR bstrKey, [out, retval] LONG* pVal);
    [propput, id(5), helpstring("property IntValue")] HRESULT IntValue([in] BSTR bstrSection, [in] BSTR bstrKey, [in] LONG newVal);
    [propget, id(6), helpstring("property BoolValue")] HRESULT BoolValue([in] BSTR bstrSection, [in] BSTR bstrKey, [out, retval] VARIANT_BOOL* pVal);
    [propput, id(6), helpstring("property BoolValue")] HRESULT BoolValue([in] BSTR bstrSection, [in] BSTR bstrKey, [in] VARIANT_BOOL newVal);
    [propget, id(7), helpstring("property Section")] HRESULT Section([in] BSTR bstrSection, [out, retval] IDispatch** pVal);
    [id(8), helpstring("method ClearSection")] HRESULT ClearSection([in] BSTR bstrSection);
    [id(9), helpstring("method Save")] HRESULT Save(void);
    [id(10), helpstring("method GetStringValue")] HRESULT GetStringValue([in] BSTR bstrSection, [in] BSTR bstrKey, [out,retval] BSTR* pbstrVal);
    [id(11), helpstring("method GetIntValue")] HRESULT GetIntValue([in] BSTR bstrSection, [in] BSTR bstrKey, [out,retval] LONG* plVal);
    [id(12), helpstring("method GetBoolValue")] HRESULT GetBoolValue([in] BSTR bstrSection, [in] BSTR bstrKey, [out,retval] VARIANT_BOOL* pboolVal);
    [id(13), helpstring("method SetStringValue")] HRESULT SetStringValue([in] BSTR bstrSection, [in] BSTR bstrKey, [in] BSTR newVal);
    [id(14), helpstring("method SetIntValue")] HRESULT SetIntValue([in] BSTR bstrSection, [in] BSTR bstrKey, [in] LONG newVal);
    [id(15), helpstring("method SetBoolValue")] HRESULT SetBoolValue([in] BSTR bstrSection, [in] BSTR bstrKey, [in] VARIANT_BOOL newVal);
};
[
    object,
    uuid(E70062FF-7718-4508-8545-3D38A13F3E9C),
    dual,
    nonextensible,
    helpstring("IWizSettingsSection Interface"),
    pointer_default(unique)
]
interface IWizSettingsSection : IDispatch{
    [propget, id(1), helpstring("property StringValue")] HRESULT StringValue([in] BSTR bstrKey, [out, retval] BSTR* pVal);
    [propget, id(2), helpstring("property IntValue")] HRESULT IntValue([in] BSTR bstrKey, [out, retval] LONG* pVal);
    [propget, id(3), helpstring("property BoolValue")] HRESULT BoolValue([in] BSTR bstrKey, [out, retval] VARIANT_BOOL* pVal);
    [id(4), helpstring("method GetStringValue")] HRESULT GetStringValue([in] BSTR bstrKey, [out,retval] BSTR* pbstrVal);
    [id(5), helpstring("method GetIntValue")] HRESULT GetIntValue([in] BSTR bstrKey, [out,retval] LONG* plVal);
    [id(6), helpstring("method GetBoolValue")] HRESULT GetBoolValue([in] BSTR bstrKey, [out,retval] VARIANT_BOOL* pboolVal);
};
[
    object,
    uuid(2A45E820-DD9D-4FF2-B66B-6CDDEDC8ABA4),
    dual,
    nonextensible,
    helpstring("IWizDeletedGUID Interface"),
    pointer_default(unique)
]
interface IWizDeletedGUID : IDispatch{
    [propget, id(1), helpstring("property GUID")] HRESULT GUID([out, retval] BSTR* pVal);
    [propget, id(2), helpstring("property Type")] HRESULT Type([out, retval] BSTR* pVal);
    [propget, id(3), helpstring("property DeletedTime")] HRESULT DeletedTime([out, retval] DATE* pVal);
    [propget, id(4), helpstring("property TypeValue")] HRESULT TypeValue([out, retval] LONG* pVal);
    [propget, id(5), helpstring("property Version")] HRESULT Version([out, retval] LONGLONG* pVal);
    [propput, id(5), helpstring("property Version")] HRESULT Version([in] LONGLONG newVal);
    [id(6), helpstring("method Delete")] HRESULT Delete(void);
};
[
    object,
    uuid(B856EEB0-1F2E-4288-B66B-4A1B8EC17708),
    dual,
    nonextensible,
    helpstring("IWizDeletedGUIDCollection Interface"),
    pointer_default(unique)
]
interface IWizDeletedGUIDCollection : IDispatch{
    [id(DISPID_NEWENUM), propget] HRESULT _NewEnum([out, retval] IUnknown** ppUnk);
    [id(DISPID_VALUE), propget] HRESULT Item([in] long Index, [out, retval] IDispatch** pVal);
    [id(1), propget] HRESULT Count([out, retval] long * pVal);
};

[
    object,
    uuid(2F5EFFF9-B2C0-44E3-9AC1-071404A30D00),
    dual,
    nonextensible,
    helpstring("IWizKMTools Interface"),
    pointer_default(unique)
]
interface IWizKMTools : IDispatch{
    [id(1), helpstring("method ZipToHtml")] HRESULT ZipToHtml([in] BSTR bstrZipFileName, [in] BSTR bstrHtmlFileName);
    [id(2), helpstring("method ZipToHtml2")] HRESULT ZipToHtml2([in] BSTR bstrZipFileName, [in] BSTR bstrFileTitleInZip, [in] BSTR bstrHtmlFileName);
    [propget, id(3), helpstring("property WindowsSearchInstalled")] HRESULT WindowsSearchInstalled([out, retval] VARIANT_BOOL* pVal);
    [propget, id(4), helpstring("property GoogleDesktopInstalled")] HRESULT GoogleDesktopInstalled([out, retval] VARIANT_BOOL* pVal);
};
[
    object,
    uuid(432C723E-9614-47F6-B380-2CB82F052A9D),
    dual,
    nonextensible,
    helpstring("IWizIndexZiwGDSPlugin Interface"),
    pointer_default(unique)
]
interface IWizIndexZiwGDSPlugin : IDispatch{
  [id(1), helpstring("method HandleFile")] HRESULT HandleFile(BSTR full_path_to_file, IDispatch *event_factory);
};
[
    object,
    uuid(B4ACEBC8-FB2F-4DD0-963A-67BFC0981F06),
    dual,
    nonextensible,
    helpstring("IWizZiwIFilter Interface"),
    pointer_default(unique)
]
interface IWizZiwIFilter : IDispatch{
};
[
    object,
    uuid(C8741532-7AC3-4EDB-87C4-E2B1C577C7AC),
    dual,
    nonextensible,
    helpstring("IWizTodoItem Interface"),
    pointer_default(unique)
]
interface IWizTodoItem : IDispatch{
    [propget, id(1), helpstring("property Text")] HRESULT Text([out, retval] BSTR* pVal);
    [propput, id(1), helpstring("property Text")] HRESULT Text([in] BSTR newVal);
    [propget, id(2), helpstring("property Prior")] HRESULT Prior([out, retval] LONG* pVal);
    [propput, id(2), helpstring("property Prior")] HRESULT Prior([in] LONG newVal);
    [propget, id(3), helpstring("property Complete")] HRESULT Complete([out, retval] LONG* pVal);
    [propput, id(3), helpstring("property Complete")] HRESULT Complete([in] LONG newVal);
    [propget, id(4), helpstring("property LinkedDocumentGUIDs")] HRESULT LinkedDocumentGUIDs([out, retval] BSTR* pVal);
    [propput, id(4), helpstring("property LinkedDocumentGUIDs")] HRESULT LinkedDocumentGUIDs([in] BSTR newVal);
    [propget, id(5), helpstring("property Blank")] HRESULT Blank([out, retval] VARIANT_BOOL* pVal);
    [propput, id(5), helpstring("property Blank")] HRESULT Blank([in] VARIANT_BOOL newVal);
    [propget, id(6), helpstring("property DateCreated")] HRESULT DateCreated([out, retval] DATE* pVal);
    [propput, id(6), helpstring("property DateCreated")] HRESULT DateCreated([in] DATE newVal);
    [propget, id(7), helpstring("property DateModified")] HRESULT DateModified([out, retval] DATE* pVal);
    [propput, id(7), helpstring("property DateModified")] HRESULT DateModified([in] DATE newVal);
    [propget, id(8), helpstring("property DateComplete")] HRESULT DateComplete([out, retval] DATE* pVal);
    [propput, id(8), helpstring("property DateComplete")] HRESULT DateComplete([in] DATE newVal);
    [propget, id(9), helpstring("property Order")] HRESULT Order([out, retval] LONGLONG* pVal);
    [propput, id(9), helpstring("property Order")] HRESULT Order([in] LONGLONG newVal);
    [propget, id(10), helpstring("property Children")] HRESULT Children([out, retval] IDispatch** pVal);
};
[
    object,
    uuid(621F668F-00AF-426B-BE84-5B92C221F571),
    dual,
    nonextensible,
    helpstring("IWizTodoItemCollection Interface"),
    pointer_default(unique)
]
interface IWizTodoItemCollection : IDispatch{
    [id(DISPID_NEWENUM), propget] HRESULT _NewEnum([out, retval] IUnknown** ppUnk);
    [id(DISPID_VALUE), propget] HRESULT Item([in] long Index, [out, retval] IDispatch** pVal);
    [id(1), propget] HRESULT Count([out, retval] long * pVal);
    [id(2), helpstring("method Add")] HRESULT Add([in] IDispatch* pTodoItemDisp);
};
[
    object,
    uuid(D5A30441-0007-4F1B-9674-286FAF90E74C),
    dual,
    nonextensible,
    helpstring("IWizEvent Interface"),
    pointer_default(unique)
]
interface IWizEvent : IDispatch{
    [propget, id(1), helpstring("property Document")] HRESULT Document([out, retval] IDispatch** pVal);
    [propget, id(2), helpstring("property GUID")] HRESULT GUID([out, retval] BSTR* pVal);
    [propget, id(3), helpstring("property Title")] HRESULT Title([out, retval] BSTR* pVal);
    [propput, id(3), helpstring("property Title")] HRESULT Title([in] BSTR newVal);
    [propget, id(4), helpstring("property Start")] HRESULT Start([out, retval] DATE* pVal);
    [propput, id(4), helpstring("property Start")] HRESULT Start([in] DATE newVal);
    [propget, id(5), helpstring("property End")] HRESULT End([out, retval] DATE* pVal);
    [propput, id(5), helpstring("property End")] HRESULT End([in] DATE newVal);
    [propget, id(6), helpstring("property Color")] HRESULT Color([out, retval] LONG* pVal);
    [propput, id(6), helpstring("property Color")] HRESULT Color([in] LONG newVal);
    [propget, id(7), helpstring("property Reminder")] HRESULT Reminder([out, retval] LONG* pVal);
    [propput, id(7), helpstring("property Reminder")] HRESULT Reminder([in] LONG newVal);
    [propget, id(8), helpstring("property Completed")] HRESULT Completed([out, retval] VARIANT_BOOL* pVal);
    [propput, id(8), helpstring("property Completed")] HRESULT Completed([in] VARIANT_BOOL newVal);
    [propget, id(9), helpstring("property Recurrence")] HRESULT Recurrence([out, retval] BSTR* pVal);
    [propput, id(9), helpstring("property Recurrence")] HRESULT Recurrence([in] BSTR newVal);
    [propget, id(10), helpstring("property EndRecurrence")] HRESULT EndRecurrence([out, retval] BSTR* pVal);
    [propput, id(10), helpstring("property EndRecurrence")] HRESULT EndRecurrence([in] BSTR newVal);
    [propget, id(11), helpstring("property RecurrenceIndex")] HRESULT RecurrenceIndex([out, retval] BSTR* pVal);
    [propput, id(11), helpstring("property RecurrenceIndex")] HRESULT RecurrenceIndex([in] BSTR newVal);
    [id(12), helpstring("method Save")] HRESULT Save(void);
    [id(13), helpstring("method Reload")] HRESULT Reload(void);
};
[
    object,
    uuid(AB1D86D6-6FE7-4CCD-BD6C-1CCF88E90CCD),
    dual,
    nonextensible,
    helpstring("IWizEventCollection Interface"),
    pointer_default(unique)
]
interface IWizEventCollection : IDispatch{
    [id(DISPID_NEWENUM), propget] HRESULT _NewEnum([out, retval] IUnknown** ppUnk);
    [id(DISPID_VALUE), propget] HRESULT Item([in] long Index, [out, retval] IDispatch** pVal);
    [id(1), propget] HRESULT Count([out, retval] long * pVal);
};
[
    object,
    uuid(74D15AA2-9E93-437C-990B-81A3D9155C06),
    dual,
    nonextensible,
    helpstring("IWizRowset Interface"),
    pointer_default(unique)
]
interface IWizRowset : IDispatch{
    [propget, id(1), helpstring("property ColumnCount")] HRESULT ColumnCount([out, retval] LONG* pVal);
    [propget, id(2), helpstring("property EOF")] HRESULT EOF([out, retval] VARIANT_BOOL* pVal);
    [id(3), helpstring("method GetFieldValue")] HRESULT GetFieldValue([in] LONG nFieldIndex, [out,retval] BSTR* pbstrFieldValue);
    [id(4), helpstring("method GetFieldValue2")] HRESULT GetFieldValue2([in] LONG nFieldIndex, [out,retval] VARIANT* pvRet);
    [id(5), helpstring("method GetFieldName")] HRESULT GetFieldName([in] LONG nFieldIndex, [out,retval] BSTR* pbstrFieldName);
    [id(6), helpstring("method GetFieldType")] HRESULT GetFieldType([in] LONG nFieldIndex, [out,retval] BSTR* pbstrFieldType);
    [id(7), helpstring("method MoveNext")] HRESULT MoveNext(void);
    [id(8), helpstring("method GetFieldValueAsStream")] HRESULT GetFieldValueAsStream([in] LONG nFieldIndex, [in] VARIANT vStream);
};
[
    object,
    uuid(AAD77AF9-FDD1-454A-A52F-7E2037EB5F8D),
    dual,
    nonextensible,
    helpstring("IWizSQLiteDatabase Interface"),
    pointer_default(unique)
]
interface IWizSQLiteDatabase : IDispatch{
    [id(1), helpstring("method Open")] HRESULT Open([in] BSTR bstrFileName);
    [id(2), helpstring("method Close")] HRESULT Close(void);
    [id(3), helpstring("method SQLQuery")] HRESULT SQLQuery([in] BSTR bstrSQL, [in] BSTR bstrOptions, [out,retval] IDispatch** ppRowsetDisp);
    [id(4), helpstring("method SQLExecute")] HRESULT SQLExecute([in] BSTR bstrSQL, [in] BSTR bstrOptions, [out,retval] LONG* pnRowsAffect);
};
[
    uuid(FF114887-DC51-4D7D-85DE-3A0D90FB7900),
    version(1.0),
    helpstring("WizKMCore 1.0 Type Library")
]
library WizKMCoreLib
{
    importlib("stdole2.tlb");
    [
        uuid(AB45C39E-7793-4DCE-8C3E-3DA52B07AD68),
        helpstring("WizDatabase Class")
    ]
    coclass WizDatabase
    {
        [default] interface IWizDatabase;
    };
    [
        uuid(95985FE1-9F9E-4F66-8A7D-4CBAEDCF5936),
        helpstring("WizTagCollection Class")
    ]
    coclass WizTagCollection
    {
        [default] interface IWizTagCollection;
    };
    [
        uuid(972019A8-D393-4EB3-B271-65E0BBBE2A2E),
        helpstring("WizTag Class")
    ]
    coclass WizTag
    {
        [default] interface IWizTag;
    };
    [
        uuid(6A94B47D-FD36-452D-952D-E4D550F10628),
        helpstring("WizDocumentCollection Class")
    ]
    coclass WizDocumentCollection
    {
        [default] interface IWizDocumentCollection;
    };
    [
        uuid(ED406C6B-C584-4029-B1A9-7FAE0A575C2B),
        helpstring("WizDocument Class")
    ]
    coclass WizDocument
    {
        [default] interface IWizDocument;
    };
    [
        uuid(C06B06EA-6E58-4761-BCB0-725EA6A3804D),
        helpstring("WizDocumentAttachmentCollection Class")
    ]
    coclass WizDocumentAttachmentCollection
    {
        [default] interface IWizDocumentAttachmentCollection;
    };
    [
        uuid(8D95BA68-C8E1-42C2-9287-CC8231C95E7F),
        helpstring("WizDocumentAttachment Class")
    ]
    coclass WizDocumentAttachment
    {
        [default] interface IWizDocumentAttachment;
    };
    [
        uuid(150F936F-26A4-474D-8DD3-811A88F0CB96),
        helpstring("WizFolderCollection Class")
    ]
    coclass WizFolderCollection
    {
        [default] interface IWizFolderCollection;
    };
    [
        uuid(2DE89910-4C9B-4252-9D30-751915121E24),
        helpstring("WizFolder Class")
    ]
    coclass WizFolder
    {
        [default] interface IWizFolder;
    };
    [
        uuid(872DD0F4-D6B6-4730-9087-757B5F66054C),
        helpstring("WizStyle Class")
    ]
    coclass WizStyle
    {
        [default] interface IWizStyle;
    };
    [
        uuid(B8B402D3-DE9E-42E5-838B-9BA0B923C16E),
        helpstring("WizStyleCollection Class")
    ]
    coclass WizStyleCollection
    {
        [default] interface IWizStyleCollection;
    };
    [
        uuid(E4DB345D-4E98-49E5-A4E8-1C923284B978),
        helpstring("WizKMProtocol Class")
    ]
    coclass WizKMProtocol
    {
        [default] interface IWizKMProtocol;
    };
    [
        uuid(8F4F05F7-D3FC-4656-8D3B-B849DA3FB6A6),
        helpstring("WizDocumentParam Class")
    ]
    coclass WizDocumentParam
    {
        [default] interface IWizDocumentParam;
    };
    [
        uuid(1B211802-4E90-440D-905E-BF6DD322B135),
        helpstring("WizDocumentParamCollection Class")
    ]
    coclass WizDocumentParamCollection
    {
        [default] interface IWizDocumentParamCollection;
    };
    [
        uuid(0DC1C265-7695-45F4-8B97-4A4B06C4DE2E),
        helpstring("WizMeta Class")
    ]
    coclass WizMeta
    {
        [default] interface IWizMeta;
    };
    [
        uuid(212C4E5C-88B0-477B-B399-A4F7FD3BC034),
        helpstring("WizMetaCollection Class")
    ]
    coclass WizMetaCollection
    {
        [default] interface IWizMetaCollection;
    };
    [
        uuid(8BC0154C-BE01-45A6-949D-A4D8E5A0D8AA),
        helpstring("WizSettings Class")
    ]
    coclass WizSettings
    {
        [default] interface IWizSettings;
    };
    [
        uuid(A6FFBEDF-C292-4CD7-83DC-CB2779D77483),
        helpstring("WizSettingsSection Class")
    ]
    coclass WizSettingsSection
    {
        [default] interface IWizSettingsSection;
    };
    [
        uuid(CDC8B878-61C3-4ECA-AC3B-7B7D8091A89B),
        helpstring("WizDeletedGUID Class")
    ]
    coclass WizDeletedGUID
    {
        [default] interface IWizDeletedGUID;
    };
    [
        uuid(D69D50B6-9A9D-4BD9-AA38-2DDD05579C26),
        helpstring("WizDeletedGUIDCollection Class")
    ]
    coclass WizDeletedGUIDCollection
    {
        [default] interface IWizDeletedGUIDCollection;
    };
    [
        uuid(0862A44D-A591-4FAF-950C-EB81BE4AC372),
        helpstring("WizKMTools Class")
    ]
    coclass WizKMTools
    {
        [default] interface IWizKMTools;
    };

    enum WizInitDocumentFlags
    { 
        [helpstring("SaveSel")]                wizUpdateDocumentSaveSel                =   0x0001,
        [helpstring("IncludeScript")]        wizUpdateDocumentIncludeScript        =   0x0002,
        [helpstring("ShowProgress")]        wizUpdateDocumentShowProgress        =   0x0004,
        [helpstring("SaveContentOnly")]        wizUpdateDocumentSaveContentOnly        =   0x0008,
        [helpstring("SaveTextOnly")]        wizUpdateDocumentSaveTextOnly        =   0x0010,
        [helpstring("DonotDownloadFile")]    wizUpdateDocumentDonotDownloadFile    =   0x0020,
        [helpstring("AllowAutoGetContent")]    wizUpdateDocumentAllowAutoGetContent    =   0x0040,
        [helpstring("DonotIncludeImages")]        wizUpdateDocumentDonotIncludeImages        =   0x0080,
        [helpstring("DonotConvertLinks")]        wizUpdateDocumentDonotConvertLinks        =   0x0100,
        [helpstring("IncludeCSSAlways")]        wizUpdateDocumentIncludeCSSAlways        =   0x0200,
        [helpstring("RunPlugins")]                wizUpdateDocumentRunPlugins                =   0x0400
    };
    enum WizDocumentGetTextFlags
    { 
        [helpstring("IncludeTitle")]            wizDocumentGetTextIncludeTitle        =   0x0001,
        [helpstring("IncludeAllFrames")]        wizDocumentGetTextIncludeAllFrames    =   0x0002,
        [helpstring("NoPromptPassword")]        wizDocumentGetTextNoPromptPassword    =   0x0004,
        [helpstring("SummaryOnly")]                wizDocumentGetTextSummaryOnly        =   0x0008,
    };
    enum WizSearchDocumentsFlags
    { 
        [helpstring("IncludeSubFolders")]        wizSearchDocumentsIncludeSubFolders    =   0x0001,
        [helpstring("FullTextSearchWindows")]            wizSearchDocumentsFullTextSearchWindows    =   0x0002,
        [helpstring("FullTextSearchGoogle")]            wizSearchDocumentsFullTextSearchGoogle    =   0x0004,
        [helpstring("FullText")]                        wizSearchDocumentsFullText                =   0x0008
    };
    enum WizFolderGetDisplayNameFlags
    { 
        [helpstring("IncludeDocumentCount")]        wizDisplayNameIncludeDocumentCount =   0x0001,
        [helpstring("IncludeDocumentCountContainsSubFolders")]        wizDisplayNameIDCContainsSubFolders=   0x0002,
    };
    enum WizGetAllFoldersDocumentCountFlags
    { 
        [helpstring("ContainsSubFolders")]        wizDocumentCountContainsSubFolders=   0x0002,
    };
    enum WizDocumentSaveToHtmlFlags
    { 
        [helpstring("UsingDisplayTemplate")]        wizDocumentToHtmlUsingDisplayTemplate    =   0x0001,
        [helpstring("UTF8")]                        wizDocumentToHtmlUTF8                    =   0x0002,
        [helpstring("UTF8BOM")]                        wizDocumentToHtmlUTF8BOM                =   0x0004,
        [helpstring("MHTFormat")]                    wizDocumentToHtmlMhtFormat                =   0x0008,
        [helpstring("DonotDownloadData")]            wizDocumentToHtmlDonotDownloadData        =   0x0010,
        [helpstring("DonotPromptPassword")]            wizDocumentToHtmlDonotPromptPassword    =   0x0020,
    };
    enum WizDatabasePasswordFlags
    { 
        [helpstring("LoginPassword")]                wizDatabasePasswordProtect                =   0x0001,
        [helpstring("ProtectWizNote")]                wizDatabasePasswordProtectWizNote        =   0x0002,
        [helpstring("ProtectWizCalendar")]            wizDatabasePasswordProtectWizCalendar    =   0x0004,
    };
    enum WizDatabaseOpenFlags
    { 
        [helpstring("NoUI")]                        wizDatabaseOpenNoUI                        =   0x0001,
        [helpstring("NoChangeAccount")]                wizDatabaseOpenNoChangeAccount            =   0x0002,
    };
    
    enum WizDatabaseSetCertFlags
    { 
        [helpstring("BackupPublicKeyToServer")]                wizDatabaseSetCertBackupPubKeyToServer        =   0x0001,
        [helpstring("BackupPrivateKeyToServer")]            wizDatabaseSetCertBackupPriKeyToServer        =   0x0002,
    };

    enum WizDocumentPrtoect
    { 
        [helpstring("ProtectNone")]                wizDocumentProtectNone        =   0,
        [helpstring("ProtectRSA")]                wizDocumentProtectRSA        =   1,
        [helpstring("ProtectAES")]                wizDocumentProtectAES        =   2,
    };
    enum WizDocumentFlags
    { 
        [helpstring("AlwaysOnTop")]                wizDocumentAlwaysOnTop        =   0x0001,
    };

    enum WizThumbImage
    {
        [helpstring("Width")]                wizThumbImageWidth        =   90,
        [helpstring("Height")]                wizThumbImageHeight        =   90,
    };

    enum WizRecentDocumentFlags
    {
        [helpstring("IncludeTopMost")]                wizRecentDocumentsIncludeTopMost        =   0x0001,
        [helpstring("IncludeRecent")]                wizRecentDocumentsIncludeRecent            =   0x0002,
    };

    [
        uuid(53829D7C-8548-4C07-A2F1-A9766B9ABA01),
        helpstring("WizIndexZiwGDSPlugin Class")
    ]
    coclass WizIndexZiwGDSPlugin
    {
        [default] interface IWizIndexZiwGDSPlugin;
    };
    [
        uuid(5ABFA0B0-FDE5-4337-B9F6-DE6CCAADC179),
        helpstring("WizZiwIFilter Class")
    ]
    coclass WizZiwIFilter
    {
        [default] interface IWizZiwIFilter;
    };
    [
        uuid(AF6ADB0C-D5CB-4B83-88E8-2CA45BE7873F),
        helpstring("WizTodoItem Class")
    ]
    coclass WizTodoItem
    {
        [default] interface IWizTodoItem;
    };
    [
        uuid(465102EA-EB28-4758-BA63-C8E177F4F3F2),
        helpstring("WizTodoItemCollection Class")
    ]
    coclass WizTodoItemCollection
    {
        [default] interface IWizTodoItemCollection;
    };
    [
        uuid(45BC1CC7-4BDD-4D17-A8F3-DC4CE48FE584),
        helpstring("WizEvent Class")
    ]
    coclass WizEvent
    {
        [default] interface IWizEvent;
    };
    [
        uuid(83287FD4-35C9-4E6D-83BC-BE341E339B0D),
        helpstring("WizEventCollection Class")
    ]
    coclass WizEventCollection
    {
        [default] interface IWizEventCollection;
    };
    [
        uuid(4D55D9D6-11EF-4816-9079-5C077D8EE864),
        helpstring("WizRowset Class")
    ]
    coclass WizRowset
    {
        [default] interface IWizRowset;
    };
    [
        uuid(CD91B37A-6B51-46D1-BC7E-06C0F9F8539D),
        helpstring("WizSQLiteDatabase Class")
    ]
    coclass WizSQLiteDatabase
    {
        [default] interface IWizSQLiteDatabase;
    };
};
```
