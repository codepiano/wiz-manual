IWizDatabase 是WizKMCore.dll包含的一个COM对象。您可以直接创建这个对象，然后打开一个用户数据库。或者通过IWizExplorerApp.Database属性来获得用户正在WizExplorer打开的数据库。

通过IWizDatabase，可以获得数据库的信息，包括各种用户信息，文档信息，文件夹，标签，样式等等。同时，也可以使用这个对象，来修改数据库，文档信息等等。

用户数据库，保存在用户数据目录下面的Data\Deault或者其它的子目录中，例如

![用户数据库目录](../img/user-database.jpg)

数据库文件夹中，包含index.db文件，该文件非常重要，用户的标签，标签组，样式，文档索引信息，附件索引信息，都保存在这个文件里面。程序会在退出的时候，自动生成一个备份文件，例如：

![数据库文件](../img/index-db.jpg)

.index.db文件是一个SQLite数据库，您可以使用相关的软件来打开这个文件，获得数据库结构以及相关的数据。

ProgID：WizKMCore.WizDatabase

文件：WizKMCore.dll

```
[
    object,
    uuid(66EDABF2-D4D0-4B63-BFFA-EB7C53A9372A),
    dual,
    nonextensible,
    helpstring("IWizDatabase Interface"),
    pointer_default(unique)
]
interface IWizDatabase : IDispatch{

    //打开一个数据库。bstrDatabasePath：指定一个数据库文件夹路径，该文件夹下面，必须有index.db文件。可以传入一个空字符串（""），如果没有默认数据库，该方法会可能显示一个对话框，让用户选择数据库。
    [id(1), helpstring("method Open")] HRESULT Open([in] BSTR bstrDatabasePath);

    //关闭数据库
    [id(2), helpstring("method Close")] HRESULT Close(void);

    //获得数据库路径。
    [propget, id(3), helpstring("property DatabasePath")] HRESULT DatabasePath([out, retval] BSTR* pVal);

    //获得数据库文件夹集合。类型为IWizFolderCollection
    [propget, id(4), helpstring("property Folders")] HRESULT Folders([out, retval] IDispatch** pVal);

    //获得根标签集合，类型为IWizTagCollection
    [propget, id(5), helpstring("property RootTags")] HRESULT RootTags([out, retval] IDispatch** pVal);

    //获得所有标签集合，类型为IWizTagCollection
    [propget, id(6), helpstring("property Tags")] HRESULT Tags([out, retval] IDispatch** pVal);

    //获得所有样式集合，类型为IWizStyleCollection
    [propget, id(7), helpstring("property Styles")] HRESULT Styles([out, retval] IDispatch** pVal);

    //获得所有附件集合，类型为IWizDocumentAttachmentCollection
    [propget, id(8), helpstring("property Attachments")] HRESULT Attachments([out, retval] IDispatch** pVal);

    //获得所有用户数据库Meta集合，类型为IWizMetaCollection
    [propget, id(9), helpstring("property Metas")] HRESULT Metas([out, retval] IDispatch** pVal);

    //通过meta的名称，获得所有名称为bstrMetaName的meta集合。bstrMetaName：meta名称；返回结果：所有名称为bstrMetaName的meta集合IWizMetaCollection
    //meta名称类似于ini文件里面的section。数据库的meta，主要用来保存用户的一些和数据库相关的设置。数据库中的meta不进行网络同步。
    [propget, id(10), helpstring("property MetasByName")] HRESULT MetasByName([in] BSTR bstrMetaName, [out, retval] IDispatch** pVal);

    //获得/设置某一个meta的值。bstrMetaName：meta名称；BSTR bstrMetaKey：meta的键值；属性值：meta的值。
    [propget, id(11), helpstring("property Meta")] HRESULT Meta([in] BSTR bstrMetaName, [in] BSTR bstrMetaKey, [out, retval] BSTR* pVal);
    [propput, id(11), helpstring("property Meta")] HRESULT Meta([in] BSTR bstrMetaName, [in] BSTR bstrMetaKey, [in] BSTR newVal);

    //获得已删除文件夹。类型为IWizFolder
    [propget, id(12), helpstring("property DeletedItemsFolder")] HRESULT DeletedItemsFolder([out, retval] IDispatch** pVal);

    //创建一个根文件夹。bstrFolderName：文件夹名称；vbSync：是否同步；返回值：创建成功的文件夹对象。类型为IWizFolder
    [id(13), helpstring("method CreateRootFolder")] HRESULT CreateRootFolder([in] BSTR bstrFolderName, [in] VARIANT_BOOL vbSync, [out,retval] IDispatch** ppNewFolderDisp);

    //创建一个根文件夹。和CreateRootFolder类似，但是可以指定自定义一个文件夹图标。bstrIconFileName：文件夹图标
    [id(14), helpstring("method CreateRootFolder2")] HRESULT CreateRootFolder2([in] BSTR bstrFolderName, [in] VARIANT_BOOL vbSync, [in] BSTR bstrIconFileName, [out,retval] IDispatch** ppNewFolderDisp);

    //创建一个根标签。bstrTagName：标签名称；bstrTagDescription：标签组描述；返回值：成功创建的标签组对象，类型为IWizTag
    [id(15), helpstring("method CreateRootTag")] HRESULT CreateRootTag([in] BSTR bstrTagName, [in] BSTR bstrTagDescription, [out,retval] IDispatch** ppNewTagDisp);

    //创建一个样式。bstrName：样式名称；bstrDescription：样式描述；nTextColor：文字颜色(RGB)；nBackColor：背景色；vbTextBold：文字是否粗体；nFlagIndex：旗帜索引，保留；返回成功创建的样式对象。
    [id(16), helpstring("method CreateStyle")] HRESULT CreateStyle([in] BSTR bstrName, [in] BSTR bstrDescription, [in] LONG nTextColor, [in] LONG nBackColor, [in] VARIANT_BOOL vbTextBold, [in] LONG nFlagIndex, [out, retval] IDispatch** ppNewStyleDisp);

    //通过标签GUID获得标签对象
    [id(18), helpstring("method TagFromGUID")] HRESULT TagFromGUID([in] BSTR bstrTagGUID, [out,retval] IDispatch** ppTagDisp);

    //通过样式GUID获得样式对象
    [id(19), helpstring("method StyleFromGUID")] HRESULT StyleFromGUID([in] BSTR bstrStyleGUID, [out,retval] IDispatch **ppStyleDisp);

    //通过文档GUID获得文档对象
    [id(20), helpstring("method DocumentFromGUID")] HRESULT DocumentFromGUID([in] BSTR bstrDocumentGUID, [out,retval] IDispatch** ppDocumentDisp);

    //通过SQL获得文档集合。您可以指定一个SQL语句查询条件，来获得文档列表。bstrWhere：sql 语句where后面的部分，例如 DOCUMENT_TITLE like '%我的文档%'；返回：IWizDocumentCollection
    [id(21), helpstring("method DocumentsFromSQL")] HRESULT DocumentsFromSQL([in] BSTR bstrSQLWhere, [out,retval] IDispatch** ppDocumentCollectionDisp);

    //获得所有具有指定标签的文档集合。pTagCollectionDisp：标签对象，类型必须是IWizTag；返回：IWizDocumentCollection
    [id(22), helpstring("method DocumentsFromTags")] HRESULT DocumentsFromTags([in] IDispatch* pTagCollectionDisp, [out,retval] IDispatch** ppDocumentCollectionDisp);

    //通过附件GUID获得附件对象
    [id(23), helpstring("method AttachmentFromGUID")] HRESULT AttachmentFromGUID([in] BSTR bstrAttachmentGUID, [out,retval] IDispatch** ppAttachmentDisp);

    //获得一个修改时间在指定时间之后的对象列表。dtTime：指定的时间；bstrObjectType：对象类型，可能的值为：tag, style, attachment, document, deleted_guid；返回值为：IWizXXXCollection
    [id(24), helpstring("method GetObjectsByTime")] HRESULT GetObjectsByTime([in] DATE dtTime, [in] BSTR bstrObjectType, [out,retval] IDispatch** ppCollectionDisp);

    //获得自从上次同步后的对象的列表。bstrObjectType：对象类型，可能的值为：tag, style, attachment, document, deleted_guid；返回值为：IWizXXXCollection
    [id(25), helpstring("method GetModifiedObjects")] HRESULT GetModifiedObjects([in] BSTR bstrObjectType, [out,retval] IDispatch** ppCollectionDisp);

    //直接创建一个Tag。需要指定所有的信息。
    [id(27), helpstring("method CreateTagEx")] HRESULT CreateTagEx([in] BSTR bstrGUID, [in] BSTR bstrGroupGUID, [in] BSTR bstrName, [in] BSTR bstrDescription, [in] DATE dtModified, [in] LONGLONG nVersion, [out,retval] IDispatch** ppRetDisp);

    //直接创建一个Style。需要指定所有的信息。
    [id(28), helpstring("method CreateStyleEx")] HRESULT CreateStyleEx([in] BSTR bstrGUID, [in] BSTR bstrName, [in] BSTR bstrDescription, [in] LONG nTextColor, [in] LONG nBackColor, [in] VARIANT_BOOL vbTextBold, [in] LONG nFlagIndex, [in] DATE dtModified, [in] LONGLONG nVersion, [out,retval] IDispatch** ppRetDisp);

    //直接创建一个Document。需要指定所有的信息。
    [id(29), helpstring("method CreateDocumentEx")] HRESULT CreateDocumentEx([in] BSTR bstrGUID, [in] BSTR bstrTitle, [in] BSTR bstrLocation, [in] BSTR bstrName, [in] BSTR bstrSEO, [in] BSTR bstrURL, [in] BSTR bstrAuthor, [in] BSTR bstrKeywords, [in] BSTR bstrType, [in] BSTR bstrOwner, [in] BSTR bstrFileType, [in] BSTR bstrStyleGUID, [in] DATE dtCreated, [in] DATE dtModified, [in] DATE dtAccessed, [in] LONG nIconIndex, [in] LONG nSync, [in] LONG nProtected, [in] LONG nReadCount, [in] LONG nAttachmentCount, [in] LONG nIndexed, [in] DATE dtInfoModified, [in] BSTR bstrInfoMD5, [in] DATE dtDataModified, [in] BSTR bstrDataMD5, [in] DATE dtParamModified, [in] BSTR bstrParamMD5, [in] VARIANT vTagGUIDs, [in] VARIANT vParams, [in] IUnknown* pDataStream, [in] LONGLONG nVersion, [out,retval] IDispatch** ppRetDisp);

    //直接创建一个Attachment。需要指定所有的信息。
    [id(30), helpstring("method CreateDocumentAttachmentEx")] HRESULT CreateDocumentAttachmentEx([in] BSTR bstrGUID, [in] BSTR bstrDocumentGUID, [in] BSTR bstrName, [in] BSTR bstrURL, [in] BSTR bstrDescription, [in] DATE dtInfoModified, [in] BSTR bstrInfoMD5, [in] DATE dtDataModified, [in] BSTR bstrDataMD5, [in] IUnknown* pDataStream, [in] LONGLONG nVersion, [out,retval] IDispatch** ppRetDisp);

    //更新一个Tag
    [id(32), helpstring("method UpdateTagEx")] HRESULT UpdateTagEx([in] BSTR bstrGUID, [in] BSTR bstrParentGUID, [in] BSTR bstrName, [in] BSTR bstrDescription, [in] DATE dtModified, [in] LONGLONG nVersion);

    //更新一个Style
    [id(33), helpstring("method UpdateStyleEx")] HRESULT UpdateStyleEx([in] BSTR bstrGUID, [in] BSTR bstrName, [in] BSTR bstrDescription, [in] LONG nTextColor, [in] LONG nBackColor, [in] VARIANT_BOOL vbTextBold, [in] LONG nFlagIndex, [in] DATE dtModified, [in] LONGLONG nVersion);

    //更新一个Document
    [id(34), helpstring("method UpdateDocumentEx")] HRESULT UpdateDocumentEx([in] BSTR bstrGUID, [in] BSTR bstrTitle, [in] BSTR bstrLocation, [in] BSTR bstrName, [in] BSTR bstrSEO, [in] BSTR bstrURL, [in] BSTR bstrAuthor, [in] BSTR bstrKeywords, [in] BSTR bstrType, [in] BSTR bstrOwner, [in] BSTR bstrFileType, [in] BSTR bstrStyleGUID, [in] DATE dtCreated, [in] DATE dtModified, [in] DATE dtAccessed, [in] LONG nIconIndex, [in] LONG nSync, [in] LONG nProtected, [in] LONG nReadCount, [in] LONG nAttachmentCount, [in] LONG nIndexed, [in] DATE dtInfoModified, [in] BSTR bstrInfoMD5, [in] DATE dtDataModified, [in] BSTR bstrDataMD5, [in] DATE dtParamModified, [in] BSTR bstrParamMD5, [in] VARIANT vTagGUIDs, [in] VARIANT vParams, [in] IUnknown* pDataStream, [in] LONGLONG nVersion, [in] long nParts);

    //更新一个Attachment
    [id(35), helpstring("method UpdateDocumentAttachmentEx")] HRESULT UpdateDocumentAttachmentEx([in] BSTR bstrGUID, [in] BSTR bstrDocumentGUID, [in] BSTR bstrName, [in] BSTR bstrURL, [in] BSTR bstrDescription, [in] DATE dtInfoModified, [in] BSTR bstrInfoMD5, [in] DATE dtDataModified, [in] BSTR bstrDataMD5, [in] IUnknown* pDataStream, [in] LONGLONG nVersion, [in] long nParts);

    //删除一个对象，而不是移动到回收站里面：bstrGUID：对象的GUID；bstrType：对象类型，可能的值为：tag, style, attachment, document, deleted_guid；dtDeleted：删除日期
    [id(36), helpstring("method DeleteObject")] HRESULT DeleteObject([in] BSTR bstrGUID, [in] BSTR bstrType, [in] DATE dtDeleted);

    //判断一个对象是否存在。bstrGUID：对象的GUID；bstrType：对象类型，可能的值为：tag, style, attachment, document, deleted_guid；返回值：对象是否存在
    [id(37), helpstring("method ObjectExists")] HRESULT ObjectExists([in] BSTR bstrGUID, [in] BSTR bstrType, [out,retval] VARIANT_BOOL* pvbExists);

    //搜索文档。bstrKeywords：搜索关键字；nFlags：搜索选项，可以是一个或者多个下面的属性：
    //wizSearchDocumentsIncludeSubFolders = 0x01，包含子文件夹
    //wizSearchDocumentsFullTextSearchWindows=0x02，使用Windows Search进行全文检索
    //wizSearchDocumentsFullTextSearchGoogle=0x04，使用Google Desktop进行全文检索
    //pFolderDisp：所在的文件夹。可以为null，表示搜索整个数据库
    //nMaxResults：返回结果最大数量
    //返回值：IWizDocumentCollection
    [id(38), helpstring("method SearchDocuments")] HRESULT SearchDocuments([in] BSTR bstrKeywords, [in] LONG nFlags, [in] IDispatch* pFolderDisp, [in] LONG nMaxResults, [out,retval] IDispatch** ppDocumentCollectionDisp);

    //清空回收站
    [id(39), helpstring("method EmptyDeletedItems")] HRESULT EmptyDeletedItems();

    //开始更新数据库。如果大批量更新数据库，可以使用BeginUpdate/EndUpdate锁定数据库，这样可以避免程序界面频繁刷新。
    [id(40), helpstring("method BeginUpdate")] HRESULT BeginUpdate();

    //停止更新
    [id(41), helpstring("method EndUpdate")] HRESULT EndUpdate();

    //获得最近修改的文档。bstrDocumentType：文档类型，例如document, note等等；nCount：返回的文档数量；返回值：IWizDocumentCollection
    [id(42), helpstring("method GetRecentDocuments")] HRESULT GetRecentDocuments([in] BSTR bstrDocumentType, [in] LONG nCount, [out,retval] IDispatch** ppDocumentCollectionDisp);

    //获得具有日历事件的文档。dtStart：开始时间；dtEnd：结束时间；返回值：IWizDocumentCollection
    [id(43), helpstring("method GetCalendarEvents")] HRESULT GetCalendarEvents([in] DATE dtStart, [in] DATE dtEnd, [out,retval] IDispatch** ppDocumentCollectionDisp);

    //通过location获得一个文件夹对象。bstrLocation：文件夹location，格式为/abc/def/；vbCreate：是否自动创建这个文件夹。
    [id(44), helpstring("method GetFolderByLocation")] HRESULT GetFolderByLocation([in] BSTR bstrLocation, [in] VARIANT_BOOL vbCreate, [out,retval] IDispatch** ppFolderDisp);

    //通过tag的名称，来获得相应的tag。bstrTagName：tag名称；vbCreate：如果tag不存在，是否创建；bstrRootTagForCreate：如果tag不存在，而且需要自动创建，则新建的tag所在的root tag名称。
    [id(45), helpstring("method GetTagByName")] HRESULT GetTagByName([in] BSTR bstrTagName, [in] VARIANT_BOOL vbCreate, [in] BSTR bstrRootTagForCreate, [out,retval] IDispatch** ppTagDisp);

    //通过root tag名称，来获得相应的root tag。bstrTagRootName：tag名称；vbCreate：是否创建
    [id(46), helpstring("method GetTagRootByName")] HRESULT GetRootTagByName([in] BSTR bstrTagRootName, [in] VARIANT_BOOL vbCreate, [out,retval] IDispatch** ppTagRootDisp);

    //获得所有的文档
    [id(47), helpstring("method GetAllDocuments")] HRESULT GetAllDocuments([out,retval] IDispatch** ppDocumentCollectionDisp);

    //获得数据库自从上次同步后，是否进行了修改。
    [id(48), helpstring("method IsModified")] HRESULT IsModified([out,retval] VARIANT_BOOL* pvbModified);

    //备份数据库索引（index.db）
    [id(49), helpstring("method BackupIndex")] HRESULT BackupIndex();

    //备份所有的数据，打包所有的数据为zip文件。bstrDestPath：备份后的文件所在的文件夹。
    [id(50), helpstring("method BackupAll")] HRESULT BackupAll([in] BSTR bstrDestPath);

    //通过一个ziw文件名（完整路径名），来获得相应的文档对象。
    [id(51), helpstring("method FileNameToDocument")] HRESULT FileNameToDocument([in] BSTR bstrFileName, [out,retval] IDispatch** ppDocumentDisp);

    //获得所有的虚拟文件夹，这是一个安全数组。元素类型为字符串
    [id(52), helpstring("method GetVirtualFolders")] HRESULT GetVirtualFolders([out,retval] VARIANT* pvVirtualFolderNames);

    //获得虚拟文件夹所包含的文档。bstrVirtualFolderName：虚拟文件夹名称
    [id(53), helpstring("method GetVirtualFolderDocuments")] HRESULT GetVirtualFolderDocuments([in] BSTR bstrVirtualFolderName, [out,retval] IDispatch** ppDocumentDisp);

    //获得虚拟文件夹的图标文件。bstrVirtualFolderName：虚拟文件夹名称；pbstrIconFileName：图标文件名
    [id(54), helpstring("method GetVirtualFolderIcon")] HRESULT GetVirtualFolderIcon([in] BSTR bstrVirtualFolderName, [out,retval] BSTR* pbstrIconFileName);

    //获得各个文件夹文档数量
    [id(55), helpstring("method GetAllFoldersDocumentCount")] HRESULT GetAllFoldersDocumentCount([in] LONG nFlags, [out,retval] VARIANT* pvDocumentCount);

    //通过文档URL来查询相同URL文档集合。返回值：IWizDocumentCollection
    [id(56), helpstring("method DocumentsFromURL")] HRESULT DocumentsFromURL([in] BSTR bstrURL, [out,retval] IDispatch** ppDocumentCollectionDisp);

    //获得文档自定义的修改时间。bstrDocumentLocation：文档所在的位置；bstrDocumentCustomID：文档自定义ID
    [id(57), helpstring("method DocumentCustomGetModified")] HRESULT DocumentCustomGetModified([in] BSTR bstrDocumentLocation, [in] BSTR bstrDocumentCustomID, [out,retval] DATE* pdtCustomModified);

    //自动更新文档。bstrDocumentLocation：文档所在的位置；bstrTitle：文档标题；bstrDocumentCustomID：文档自定义ID；bstrDocumentURL：文档URL；dtCustomModified：文档自定义修改时间；bstrHtmlFileName：文档html文件；vbAllowOverwrite：是否允许覆盖已经存在的文档（根据bstrDocumentCustomID查询是否已经存在相应的文档）；nUpdateDocumentFlags：更新文档选项。
    [id(58), helpstring("method DocumentCustomUpdate")] HRESULT DocumentCustomUpdate([in] BSTR bstrDocumentLocation, [in] BSTR bstrTitle, [in] BSTR bstrDocumentCustomID, [in] BSTR bstrDocumentURL, [in] DATE dtCustomModified, [in] BSTR bstrHtmlFileName, [in] VARIANT_BOOL vbAllowOverwrite, [in] LONG nUpdateDocumentFlags, [out,retval] IDispatch** ppDocumentDisp);

    //设置是否启用文档全文索引功能

    [id(59), helpstring("method EnableDocumentIndexing")] HRESULT EnableDocumentIndexing([in] VARIANT_BOOL vbEnable);

    //导出数据库。已经无效
    [id(60), helpstring("method DumpDatabase")] HRESULT DumpDatabase([in] LONG nFlags, [out,retval] VARIANT_BOOL* pvbRet);

    //重建数据库，已经无效
    [id(61), helpstring("method RebuildDatabase")] HRESULT RebuildDatabase([in] LONG nFlags, [out,retval] VARIANT_BOOL* pvbRet);

    //获得todo list对象的文档
    [id(62), helpstring("method GetTodoDocument")] HRESULT GetTodoDocument([in] DATE dtDate, [out,retval] IDispatch** ppDocumentDisp);

    //获得日历中的事件
    [id(63), helpstring("method GetCalendarEvents2")] HRESULT GetCalendarEvents2([in] DATE dtStart, [in] DATE dtEnd, [out,retval] IDispatch** ppEventCollectionDisp);

    //获得系统标签组。（通过标签组的名称，会进行本地化）
    [id(64), helpstring("method GetKnownRootTagName")] HRESULT GetKnownRootTagName([in] BSTR bstrTagRootName, [out,retval] BSTR* pbstrLocalRootTagName);

    //执行一个SQL查询，返回一个IWizRowset数据集
    [id(65), helpstring("method SQLQuery")] HRESULT SQLQuery([in] BSTR bstrSQL, [in] BSTR bstrOptions, [out,retval] IDispatch** ppRowsetDisp);

    //执行一个SQL更新操作（insert update delete）
    [id(66), helpstring("method SQLExecute")] HRESULT SQLExecute([in] BSTR bstrSQL, [in] BSTR bstrOptions, [out,retval] LONG* pnRowsAffect);

    //获得标签对应的文档数量
    [id(67), helpstring("method GetAllTagsDocumentCount")] HRESULT GetAllTagsDocumentCount([in] LONG nFlags, [out,retval] VARIANT* pvDocumentCount);

    //直接添加一个ziw文件作为文档。这个ziw文件必须在数据库目录下面，而且不在数据库中索引中存在。
    [id(68), helpstring("method AddZiwFile")] HRESULT AddZiwFile([in] LONG nFlags, [in] BSTR bstrZiwFileName, [out,retval] IDispatch** ppDocumentDisp);

    //打开数据库
    [id(69), helpstring("method Open2")] HRESULT Open2([in] BSTR bstrDatabasePath, [in] BSTR bstrPassword, [in] LONG nFlags, [out,retval] BSTR* pbstrErrorMessage);

    //更改数据库密码
    [id(70), helpstring("method ChangePassword")] HRESULT ChangePassword([in] BSTR bstrOldPassword, [in] BSTR bstrPassword);

    //密码属性
    [propget, id(71), helpstring("property PasswordFlags")] HRESULT PasswordFlags([out, retval] LONG* pVal);
    [propput, id(71), helpstring("property PasswordFlags")] HRESULT PasswordFlags([in] LONG newVal);

    //当前数据库对应的wizkm的用户名
    [propget, id(72), helpstring("property UserName")] HRESULT UserName([out, retval] BSTR* pVal);
    [propput, id(72), helpstring("property UserName")] HRESULT UserName([in] BSTR newVal);

    //获得密码，该密码被加密
    [id(73), helpstring("method GetEncryptedPassword")] HRESULT GetEncryptedPassword([out,retval] BSTR* pbstrPassword);

    //获得用户输入的密码
    [id(74), helpstring("method GetUserPassword")] HRESULT GetUserPassword([out,retval] BSTR* pbstrPassword);

    //设置证书到数据库
    [id(75), helpstring("method SerCert")] HRESULT SetCert([in] BSTR bstrN, [in] BSTR bstre, [in] BSTR bstrEcnryptedd, [in] BSTR bstrHint, [in] LONG nFlags, [in] LONGLONG nWindowHandle);

    //获得证书的N
    [id(76), helpstring("method GetCertN")] HRESULT GetCertN([out,retval] BSTR* pbstrN);

    //获得证书的e
    [id(77), helpstring("method GetCerte")] HRESULT GetCerte([out,retval] BSTR* pbstre);

    //获得证书的d，被加密
    [id(78), helpstring("method GetEncryptedCertd")] HRESULT GetEncryptedCertd([out,retval] BSTR* pbstrEncrypted_d);

    //获得证书密码的提示信息
    [id(79), helpstring("method GetCertHint")] HRESULT GetCertHint([out,retval] BSTR* pbstrHint);

    //从服务器初始化证书（如果用户曾经将证书备份到服务器）
    [id(80), helpstring("method InitCertFromServer")] HRESULT InitCertFromServer([in] LONGLONG nWindowHandle);

    //证书密码
    [propget, id(81), helpstring("property CertPassword")] HRESULT CertPassword([out, retval] BSTR* pVal);
    [propput, id(81), helpstring("property CertPassword")] HRESULT CertPassword([in] BSTR newVal);

    //获得所有的文档标题
    [propget, id(82), helpstring("property AllDocumentsTitle")] HRESULT AllDocumentsTitle([out, retval] BSTR* pVal);

    //通过文档标题查询文档
    [id(83), helpstring("method DocumentsFromTitle")] HRESULT DocumentsFromTitle([in] BSTR bstrTitle, [in] LONG nFlags, [out,retval] IDispatch** ppDocumentCollectionDisp);

    //通过标签获得笔记，包含子标签对应的笔记
    [id(84), helpstring("method DocumentsFromTagWithChildren")] HRESULT DocumentsFromTagWithChildren([in] IDispatch* pRootTagDisp, [out,retval] IDispatch** ppDocumentCollectionDisp);

    //获得最近的笔记
    [id(85), helpstring("method GetRecentDocuments2")] HRESULT GetRecentDocuments2([in] LONG nFlags, [in] BSTR bstrDocumentType, [in] LONG nCount, [out,retval] IDispatch** ppDocumentCollectionDisp);

    //创建任务列表
    [id(86), helpstring("method CreateTodo2Document")] HRESULT CreateTodo2Document([in] BSTR bstrLocation, [in] BSTR bstrTitle, [out,retval] IDispatch** ppDocumentDisp);

    //移动完成的任务列表到已完成
    [id(87), helpstring("method MoveCompletedTodoItems")] HRESULT MoveCompletedTodoItems([in] IDispatch* pSrcTodoDocumentDisp, [in] IDispatch* pDestTodoDocumentDisp);

    //获得任务列表对应的笔记对象
    [id(88), helpstring("method GetTodo2Documents")] HRESULT GetTodo2Documents([in] BSTR bstrLocation, [out, retval] IDispatch** ppDocumentCollectionDisp);

    //获得一个文件夹下面，具有某一个些标签的笔记
    [id(89), helpstring("method DocumentsFromTagsInFolder")] HRESULT DocumentsFromTagsInFolder([in] IDispatch* pFolderDisp, [in] IDispatch* pTagsCollDisp, [out, retval] IDispatch** ppDocumentCollectionDisp);

    //获得一个文件夹下面，具有某一个些标签的笔记
    [id(90), helpstring("method DocumentsFromStyleInFolder")] HRESULT DocumentsFromStyleInFolder([in] IDispatch* pFolderDisp, [in] IDispatch* pStyleDisp, [out, retval] IDispatch** ppDocumentCollectionDisp);

    //通过SQL语句获得笔记，注意：只能是where后面部分
    [id(91), helpstring("method DocumentsFromSQL2")] HRESULT DocumentsFromSQL2([in] BSTR bstrSQLWhere, [in] BSTR bstrOptions, [out,retval] IDispatch** ppDocumentCollectionDisp);

    //获得所有的文件夹
    [propget, id(92), helpstring("property AllLocations")] HRESULT AllLocations([out, retval] VARIANT* pVal);

    //是否自定义排序
    [propget, id(93), helpstring("property IsCustomSorted")] HRESULT IsCustomSorted([out, retval] VARIANT_BOOL* pVal);

    //获得一个文件夹下面，具有某一个些标签的笔记
    [id(94), helpstring("method DocumentsFromTagsInFolder2")] HRESULT DocumentsFromTagsInFolder2([in] VARIANT_BOOL vbIncludeChildTags, [in] VARIANT_BOOL vbTagsAnd, [in] IDispatch* pFolderDisp, [in] IDispatch* pTagsCollDisp, [out, retval] IDispatch** ppDocumentCollectionDisp);

    //注册日志输出窗口
    [id(95), helpstring("method RegisterLogWindow")] HRESULT RegisterLogWindow([in] LONGLONG hwndWindow);

    //修复笔记索引
    [id(96), helpstring("method RepairIndex")] HRESULT RepairIndex([in] BSTR bstrDestFileName);

    //获得一个Meta值
    [id(97), helpstring("method GetMeta")] HRESULT GetMeta([in] BSTR bstrMetaName, [in] BSTR bstrMetaKey, [out,retval] BSTR* pbstrVal);

    //设置一个Meta值，用于保存某些和数据库相关的设置
    [id(98), helpstring("method SetMeta")] HRESULT SetMeta([in] BSTR bstrMetaName, [in] BSTR bstrMetaKey, [in] BSTR newVal);

    //打开一个数据库，可以指定打开一个群组数据库
    [id(99), helpstring("method Open3")] HRESULT Open3([in] BSTR bstrDatabasePath, [in] BSTR bstrPassword, [in] BSTR bstrKbGUID, [in] LONG nFlags, [out,retval] BSTR* pbstrErrorMessage);

    //获得群组的KbGUID
    [propget, id(100), helpstring("property KbGUID")] HRESULT KbGUID([out, retval] BSTR* pVal);

    //获得需要下载的对象的列表
    [id(101), helpstring("method GetObjectsNeedToBeDownloaded")] HRESULT GetObjectsNeedToBeDownloaded([in] BSTR bstrObjectType, [in] BSTR bstrExtCondition, [out,retval] IDispatch** ppCollectionDisp);

    //判断是否是一个群组数据库
    [propget, id(102), helpstring("property IsGroupDatabase")] HRESULT IsGroupDatabase([out, retval] VARIANT_BOOL* pVal);

    //设置一个对象是否需要从服务器下载
    [id(103), helpstring("method SetObjectDownloaded")] HRESULT SetObjectDownloaded([in] BSTR bstrGUID, [in] BSTR bstrObjectType, [in] VARIANT_BOOL vbDownloaded);

    //获得常用的标签
    [id(104), helpstring("method GetCommonUsedTags")] HRESULT GetCommonUsedTags([in] LONG count, [out,retval] IDispatch** ppDisp);

    //获得所有的标签名称
    [id(105), helpstring("method GetAllTagsName")] HRESULT GetAllTagsName([out,retval] BSTR* pbstrAllTagsName);

    //获得用户在群组中的角色
    [propget, id(106), helpstring("property UserGroup")] HRESULT UserGroup([out, retval] LONG* pVal);

    //获得群组未读笔记数量
    [propget, id(107), helpstring("property UnreadCount")] HRESULT UnreadCount([out, retval] LONG* pVal);

    //标记群组中所有的笔记为已读
    [id(108), helpstring("method MarkAllRead")] HRESULT MarkAllRead();

    //获得所有的标签对应的笔记数量
    [id(109), helpstring("method GetAllTagsDocumentCount2")] HRESULT GetAllTagsDocumentCount2([in] IDispatch* pFolderDisp, [in] LONG nFlags, [out,retval] VARIANT* pvDocumentCount);

    //获得群组名称
    [propget, id(110), helpstring("property GroupName")] HRESULT GroupName([out,retval] BSTR* pVal);

    //创建Wiz内部对象
    [id(111), helpstring("method CreateWizObject")] HRESULT CreateWizObject([in] BSTR bstrObjectName, [out,retval] IDispatch** ppVal);

};
```
