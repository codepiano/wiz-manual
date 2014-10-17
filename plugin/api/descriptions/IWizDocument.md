IWizDocument是WizKMCore.dll包含的一个COM对象。文档对象必须隶属于一个数据库，因此您不能直接创建这个对象，而是需要通过IWizDatabase对象来获得数据库中的文档对象。
通过IWizDatabase.GetAllDocuments，获得IWizFolder.Documents等等方法，可以获得数据库的文档信息。

每一个IWizDocument对象，都在索引数据库(index.db)里面，包含一条记录，记录了文档的元数据，同时，还有一个磁盘文件(\*.ziw)，保存了文档的数据。
ziw是一个zip文件，即使没有安装Wiz，您也可以用压缩软件来打开ziw文件，查看完整的html内容。因为使用了zip方式压缩，不但可以减少磁盘占用，还可以保持磁盘整洁，减少磁盘碎片，而且，即使以后您不再使用Wiz文件，您的文档也不会因为没有安装Wiz而无法打开。

Wiz提供了WizViewer.exe，专门用来直接查看ziw文件。在我的电脑里面，直接双击一个ziw文件，就可以用WizViewer.exe来打开文件。另外，Wiz提供了Windows Search的IFilter插件和Google Desktop的插件，可以用来对Wiz来建立索引，提供桌面全文检索。

请不要直接删除ziw文件，因为这样会破坏数据完整性。

文档可以包含任意数量的附件。每一个附件，都对应于一个磁盘文件。附件保存在文档数据文件夹下面的XXX_Attachments文件夹内。

文档可以包含任意数量的自定义参数，自定义参数，主要是用来扩展Wiz的功能。

文件：WizKMCore.dll

```
[
    object,
    uuid(5B78A65E-6EE8-41C7-B1DA-8ECBF9D917B0),
    dual,
    nonextensible,
    helpstring("IWizDocument Interface"),
    pointer_default(unique)
]
interface IWizDocument : IDispatch{

    //获得文档的GUID
    [propget, id(1), helpstring("property GUID")] HRESULT GUID([out, retval] BSTR* pVal);

    //获得/设置文档的标题
    [propget, id(2), helpstring("property Title")] HRESULT Title([out, retval] BSTR* pVal);
    [propput, id(2), helpstring("property Title")] HRESULT Title([in] BSTR newVal);

    //获得/设置文档的作者
    [propget, id(3), helpstring("property Author")] HRESULT Author([out, retval] BSTR* pVal);
    [propput, id(3), helpstring("property Author")] HRESULT Author([in] BSTR newVal);

    //获得/设置文档的关键字
    [propget, id(4), helpstring("property Keywords")] HRESULT Keywords([out, retval] BSTR* pVal);
    [propput, id(4), helpstring("property Keywords")] HRESULT Keywords([in] BSTR newVal);

    //获得/设置文档的标签，类型为IWizTagCollection
    [propget, id(5), helpstring("property Tags")] HRESULT Tags([out, retval] IDispatch** pVal);
    [propput, id(5), helpstring("property Tags")] HRESULT Tags([in] IDispatch* newVal);

    //获得/设置文档的名称（对应的ziw的文件名，不包含路径）
    [propget, id(6), helpstring("property Name")] HRESULT Name([out, retval] BSTR* pVal);
    [propput, id(6), helpstring("property Name")] HRESULT Name([in] BSTR newVal);

    //获得文档在数据库中的路径。格式为/abx/def/
    [propget, id(7), helpstring("property Location")] HRESULT Location([out, retval] BSTR* pVal);

    //获得文档数据的完整路径名（ziw的完整路径名）
    [propget, id(8), helpstring("property FileName")] HRESULT FileName([out, retval] BSTR* pVal);

    //获得/设置文档在网络上面的url名称。如果文档不公开，这个数据无用
    [propget, id(9), helpstring("property SEO")] HRESULT SEO([out, retval] BSTR* pVal);
    [propput, id(9), helpstring("property SEO")] HRESULT SEO([in] BSTR newVal);

    //获得/设置文档的原始网页URL，或者磁盘文件名（导入的文件）
    [propget, id(10), helpstring("property URL")] HRESULT URL([out, retval] BSTR* pVal);
    [propput, id(10), helpstring("property URL")] HRESULT URL([in] BSTR newVal);

    //获得/设置文档的类型，例如document，note，journal，contact等等
    [propget, id(11), helpstring("property Type")] HRESULT Type([out, retval] BSTR* pVal);
    [propput, id(11), helpstring("property Type")] HRESULT Type([in] BSTR newVal);

    //获得/设置文档的所有者，默认是建立文档所在的电脑名称
    [propget, id(12), helpstring("property Owner")] HRESULT Owner([out, retval] BSTR* pVal);
    [propput, id(12), helpstring("property Owner")] HRESULT Owner([in] BSTR newVal);

    //获得/设置文档的文件类型，例如html，doc，jpg等等
    [propget, id(13), helpstring("property FileType")] HRESULT FileType([out, retval] BSTR* pVal);
    [propput, id(13), helpstring("property FileType")] HRESULT FileType([in] BSTR newVal);

    //获得/设置文档的样式，类型为IWizStyle
    [propget, id(14), helpstring("property Style")] HRESULT Style([out, retval] IDispatch** pVal);
    [propput, id(14), helpstring("property Style")] HRESULT Style([in] IDispatch* newVal);

    //获得/设置文档的图标索引，暂不支持
    [propget, id(15), helpstring("property IconIndex")] HRESULT IconIndex([out, retval] LONG* pVal);
    [propput, id(15), helpstring("property IconIndex")] HRESULT IconIndex([in] LONG newVal);

    //获得/设置文档是否进行同步，暂不支持
    [propget, id(16), helpstring("property Sync")] HRESULT Sync([out, retval] VARIANT_BOOL* pVal);
    [propput, id(16), helpstring("property Sync")] HRESULT Sync([in] VARIANT_BOOL newVal);

    //获得/设置文档的保护类型，暂不支持
    [propget, id(17), helpstring("property Protect")] HRESULT Protect([out, retval] LONG* pVal);
    [propput, id(17), helpstring("property Protect")] HRESULT Protect([in] LONG newVal);

    //获得/设置文档的阅读次数
    [propget, id(18), helpstring("property ReadCount")] HRESULT ReadCount([out, retval] LONG* pVal);
    [propput, id(18), helpstring("property ReadCount")] HRESULT ReadCount([in] LONG newVal);

    //获得文档的附件数量
    [propget, id(19), helpstring("property AtachmentCount")] HRESULT AttachmentCount([out, retval] LONG* pVal);

    //获得/设置文档是否被索引了，暂不支持
    [propget, id(20), helpstring("property Indexed")] HRESULT Indexed([out, retval] VARIANT_BOOL* pVal);
    [propput, id(20), helpstring("property Indexed")] HRESULT Indexed([in] VARIANT_BOOL newVal);

    //获得/设置文档的创建时间
    [propget, id(21), helpstring("property DateCreated")] HRESULT DateCreated([out, retval] DATE* pVal);
    [propput, id(21), helpstring("property DateCreated")] HRESULT DateCreated([in] DATE newVal);

    //获得/设置文档的修改时间
    [propget, id(22), helpstring("property DateModified")] HRESULT DateModified([out, retval] DATE* pVal);
    [propput, id(22), helpstring("property DateModified")] HRESULT DateModified([in] DATE newVal);

    //获得/设置文档的访问时间
    [propget, id(23), helpstring("property DateAccessed")] HRESULT DateAccessed([out, retval] DATE* pVal);
    [propput, id(23), helpstring("property DateAccessed")] HRESULT DateAccessed([in] DATE newVal);

    //获得文档的信息被修改的时间
    [propget, id(24), helpstring("property InfoDateModified")] HRESULT InfoDateModified([out, retval] DATE* pVal);

    //获得文档的信息md5值
    [propget, id(25), helpstring("property InfoMD5")] HRESULT InfoMD5([out, retval] BSTR* pVal);

    //获得文档的数据修改时间
    [propget, id(26), helpstring("property DataDateModified")] HRESULT DataDateModified([out, retval] DATE* pVal);

    //获得文档的数据md5值
    [propget, id(27), helpstring("property DataMD5")] HRESULT DataMD5([out, retval] BSTR* pVal);

    //获得文档的参数修改时间
    [propget, id(28), helpstring("property ParamDateModified")] HRESULT ParamDateModified([out, retval] DATE* pVal);

    //获得文档的参数md5值
    [propget, id(29), helpstring("property ParamMD5")] HRESULT ParamMD5([out, retval] BSTR* pVal);

    //获得文档的所有附件，类型为IWizDocumentAttachmentCollection
    [propget, id(30), helpstring("property Attachments")] HRESULT Attachments([out, retval] IDispatch** pVal);

    //获得/设置文档的参数。bstrParamName：参数名称，忽略大小写
    [propget, id(31), helpstring("property ParamValue")] HRESULT ParamValue([in] BSTR bstrParamName, [out, retval] BSTR* pVal);
    [propput, id(31), helpstring("property ParamValue")] HRESULT ParamValue([in] BSTR bstrParamName, [in] BSTR newVal);

    //获得文档的所有参数。类型为IWizDocumentParamCollection
    [propget, id(32), helpstring("property Params")] HRESULT Params([out, retval] IDispatch** pVal);

    //获得文档所在的文件夹，类型为IWizFolder
    [propget, id(33), helpstring("property Parent")] HRESULT Parent([out, retval] IDispatch** pVal);

    //获得/设置文档的样式GUID
    [propget, id(34), helpstring("property StyleGUID")] HRESULT StyleGUID([out, retval] BSTR* pVal);
    [propput, id(34), helpstring("property StyleGUID")] HRESULT StyleGUID([in] BSTR newVal);

    //获得/设置文档的版本，用于同步，建议不要修改，否则可能导致同步错误
    [propget, id(35), helpstring("property Version")] HRESULT Version([out, retval] LONGLONG* pVal);
    [propput, id(35), helpstring("property Version")] HRESULT Version([in] LONGLONG newVal);

    //获得/设置文档的标签文字，每一个标签之间，通过英文半角分号(;)分割。
    [propget, id(36), helpstring("property TagsText")] HRESULT TagsText([out, retval] BSTR* pVal);
    [propput, id(36), helpstring("property TagsText")] HRESULT TagsText([in] BSTR newVal);

    //获得文档的所在的数据库
    [propget, id(37), helpstring("property Database")] HRESULT Database([out, retval] IDispatch** pVal);

    //获得文档数据所在的文件夹路径，例如：D:\My Documents\My Knowledge\Data\Default\新闻\
    [propget, id(38), helpstring("property FilePath")] HRESULT FilePath([out, retval] BSTR* pVal);

    //获得文档附件所在的文件夹路径，例如：D:\My Documents\My Knowledge\Data\Default\新闻\abc_Attachments\
    [propget, id(39), helpstring("property AttachmentsFilePath")] HRESULT AttachmentsFilePath([out, retval] BSTR* pVal);

    //添加一个标签。pTagDisp：标签对象，类型为IWizTag
    [id(40), helpstring("method AddTag")] HRESULT AddTag([in] IDispatch* pTagDisp);

    //删除一个标签。pTagDisp：标签对象，类型为IWizTag
    [id(41), helpstring("method RemoveTag")] HRESULT RemoveTag([in] IDispatch* pTagDisp);

    //删除全部的文档标签。
    [id(42), helpstring("method RemoveAllTags")] HRESULT RemoveAllTags();

    //删除文档
    [id(43), helpstring("method Delete")] HRESULT Delete(void);

    //重新从数据库里面获得文档的信息。
    [id(44), helpstring("method Reload")] HRESULT Reload(void);

    //添加一个附件。bstrFileName：磁盘文件名；返回值：成功添加的附件的对象，类型为IWizDocumentAttachment
    [id(45), helpstring("method AddAttachment")] HRESULT AddAttachment([in] BSTR bstrFileName, [out,retval] IDispatch** ppNewAttachmentDisp);

    //保存为html文件。bstrHtmlFileName：保存后的html文件名；nFlags：选项，可能的值为：wizDocumentToHtmlUsingDisplayTemplate=0x01，使用现实模板。
    [id(46), helpstring("method SaveToHtml")] HRESULT SaveToHtml([in] BSTR bstrHtmlFileName, [in] LONG nFlags);

    //移动到目标文件夹。pDestFolderDisp：目标文件夹，类型为IWizFolder
    [id(47), helpstring("method MoveTo")] HRESULT MoveTo([in] IDispatch* pDestFolderDisp);

    //复制到到目标文件夹。pDestFolderDisp：目标文件夹，类型为IWizFolder；返回值：复制后的新的文档对象，类型为IWizDocument
    [id(48), helpstring("method CopyTo")] HRESULT CopyTo([in] IDispatch* pDestFolderDisp, [out,retval] IDispatch** ppNewDocumentDisp);

    //判断文档是否在一个文件夹内。pFolderDisp：目标文件夹。注意：只要文档在文件夹下面（可以多级），就返回true。例如用来判断一个文档是否在回收站内。
    [id(49), helpstring("method IsIn")] HRESULT IsIn([in] IDispatch* pFolderDisp, [out,retval] VARIANT_BOOL* pvbRet);

    //更改文档数据，通过一个html文件来更新文档。
    //wizUpdateDocumentSaveSel                = 0x0001    保存选中部分，仅仅针对UpdateDocument2有效
    //wizUpdateDocumentIncludeScript        = 0x0002    包含html里面的脚本
    //wizUpdateDocumentShowProgress        = 0x0004    显示进度
    //wizUpdateDocumentSaveContentOnly        = 0x0008   只保存正文 
    //wizUpdateDocumentSaveTextOnly        = 0x0010    只保存文字内容，并且为纯文本
    //wizUpdateDocumentDonotDownloadFile    = 0x0020    不从网络下载html里面的资源
    //wizUpdateDocumentAllowAutoGetContent    = 0x0040    如果只保存正文，允许使用自动获得正文方式
    [id(50), helpstring("method UpdateDocument")] HRESULT UpdateDocument([in] BSTR bstrHtmlFileName, [in] LONG nFlags);

    //更改文档数据，通过一个IHTMLDocument2对象来更新
    [id(51), helpstring("method UpdateDocument2")] HRESULT UpdateDocument2([in] IDispatch* pIHTMLDocument2Disp, [in] LONG nFlags);

    //更改文档数据，通过HTML文字内容来更新
    [id(52), helpstring("method UpdateDocument3")] HRESULT UpdateDocument3([in] BSTR bstrHtml, [in] LONG nFlags);

    //更改文档数据，通过HTML文字内容以及该HTML的URL来更新
    [id(53), helpstring("method UpdateDocument4")] HRESULT UpdateDocument4([in] BSTR bstrHtml, [in] BSTR bstrURL, [in] LONG nFlags);

    //更改文档数据，通过一个HTML文件名来更新。
    [id(54), helpstring("method UpdateDocument5")] HRESULT UpdateDocument5([in] BSTR bstrHtmlFileName);

    //更改文档数据，通过一个HTML文件名和对应的URL来更新。
    [id(55), helpstring("method UpdateDocument6")] HRESULT UpdateDocument6([in] BSTR bstrHtmlFileName, [in] BSTR bstrURL, [in] LONG nFlags);

    //比较两个文档，保留
    [id(56), helpstring("method CompareTo")] HRESULT CompareTo([in] IDispatch* pDocumentDisp, [in] LONG nCompareBy, [out,retval] LONG* pnRet);

    //获得文档的text内容
    [id(57), helpstring("method GetText")] HRESULT GetText([in] UINT nFlags, [out, retval] BSTR* pbstrText);

    //获得文档的html内容
    [id(58), helpstring("method GetHtml")] HRESULT GetHtml([out, retval] BSTR* pbstrHtml);

    //开始更新文档的参数
    [id(59), helpstring("method BeginUpdateParams")] HRESULT BeginUpdateParams();

    //结束更新文档的参数
    [id(60), helpstring("method EndUpdateParams")] HRESULT EndUpdateParams();

    //添加文档到日历。dtStart：开始时间；dtEnd：结束时间
    [id(61), helpstring("method AddToCalendar")] HRESULT AddToCalendar([in] DATE dtStart, [in] DATE dtEnd, [in] BSTR bstrExtInfo);

    //从日历中删除文档
    [id(62), helpstring("method RemoveFromCalendar")] HRESULT RemoveFromCalendar(void);

    //更改文档标题和文件名
    [id(63), helpstring("method ChangeTitleAndFileName")] HRESULT ChangeTitleAndFileName([in] BSTR bstrTitle);

    //获得文档的图标文件名
    [id(64), helpstring("method GetIconFileName")] HRESULT GetIconFileName([out,retval] BSTR* pbstrIconFileName);
};
    //获得todo list数据（如果文档包含todo list数据的话）

    [propget, id(65), helpstring("property TodoItems")] HRESULT TodoItems([out, retval] IDispatch** pVal);
    [propput, id(65), helpstring("property TodoItems")] HRESULT TodoItems([in] IDispatch* newVal);

    //获得事件数据（如果文档包含事件数据的话）
    [propget, id(66), helpstring("property Event")] HRESULT Event([out, retval] IDispatch** pVal);
    [propput, id(66), helpstring("property Event")] HRESULT Event([in] IDispatch* newVal);

    //添加文档到日历中，可以设置重复参数
    [id(67), helpstring("method AddToCalendar2")] HRESULT AddToCalendar2([in] DATE dtStart, [in] DATE dtEnd, [in] BSTR bstrRecurrence, [in] BSTR bstrEndRecurrence, [in] BSTR bstrExtInfo);

    //设置标签文字，可以用逗号分隔多个标签。
    [id(68), helpstring("method SetTagsText2")] HRESULT SetTagsText2([in] BSTR bstrTagsText, [in] BSTR bstrNewTagGroupName);

    //彻底删除文档
    [id(69), helpstring("method PermanentlyDelete")] HRESULT PermanentlyDelete(void);

    //获得文档是否被加密。
    [id(70), helpstring("method GetProtectedByData")] HRESULT GetProtectedByData([out,retval] LONG* pnProtected);

    //自动加密文档（通过用户的设置加密，如果需要加密，则加密数据忙否则不加密）
    [id(71), helpstring("method AutoEncrypt")] HRESULT AutoEncrypt([in] VARIANT_BOOL vbDecrypt);

    //笔记附属属性，内部使用

    [propget, id(72), helpstring("property Flags")] HRESULT Flags([out, retval] LONG* pVal);
    [propput, id(72), helpstring("property Flags")] HRESULT Flags([in] LONG newVal);

    //设置/获取笔记置顶属性
    [propget, id(73), helpstring("property AlwaysOnTop")] HRESULT AlwaysOnTop([out, retval] VARIANT_BOOL* pVal);
    [propput, id(73), helpstring("property AlwaysOnTop")] HRESULT AlwaysOnTop([in] VARIANT_BOOL newVal);

    //设置/获取笔记评价属性
    [propget, id(74), helpstring("property Rate")] HRESULT Rate([out, retval] LONG* pVal);
    [propput, id(74), helpstring("property Rate")] HRESULT Rate([in] LONG newVal);

    //设置数据最后修改时间

    [id(78), helpstring("method SetDataDateModified")] HRESULT SetDataDateModified([in] DATE dtDataModified);

    //获取/设置参数值
    [id(80), helpstring("method GetParamValue")] HRESULT GetParamValue([in] BSTR bstrParamName, [out,retval] BSTR* pbstrVal);
    [id(81), helpstring("method SetParamValue")] HRESULT SetParamValue([in] BSTR bstrParamName, [in] BSTR newVal);

    //设置/获取笔记数据是否已经下载到本地
    [propget, id(82), helpstring("property Downloaded")] HRESULT Downloaded([out, retval] VARIANT_BOOL* pVal);
    [propput, id(82), helpstring("property Downloaded")] HRESULT Downloaded([in] VARIANT_BOOL newVal);

    //检查笔记数据是否已经下载到本地，如果没有，则自动下载。返回true或者false。如果失败则返回false
    [id(83), helpstring("method CheckDocumentData")] HRESULT CheckDocumentData([in] LONGLONG hwndParent, [out,retval] VARIANT_BOOL* vbRet);

    //设置/获得笔记数据，类型是IStream，脚本语言无法使用
    [id(84), helpstring("method SetData")] HRESULT SetData([in] IUnknown* pData);
    [id(85), helpstring("method GetData")] HRESULT GetData([in] IUnknown** ppData);

    //获取摘要文字
    [propget, id(86), helpstring("property AbstractText")] HRESULT AbstractText([out, retval] BSTR* pVal);

    //获取摘要图片，类型是HBITMPA，脚本语言无法使用
    [propget, id(87), helpstring("property AbstractImage")] HRESULT AbstractImage([out, retval] LONGLONG* pVal);

    //获取摘要图片，类型是HBITMPA，脚本语言无法使用
    [propget, id(88), helpstring("property AbstractImageOwner")] HRESULT AbstractImageOwner([out, retval] LONGLONG* pVal);

    //获取笔记是否可以被编辑（用于群组笔记）
    [propget, id(89), helpstring("property CanEdit")] HRESULT CanEdit([out, retval] VARIANT_BOOL* pVal);

    //获取/设置笔记的服务器版本号，内部使用
    [propget, id(90), helpstring("property ServerVersion")] HRESULT ServerVersion([out, retval] LONGLONG* pVal);
    [propput, id(90), helpstring("property ServerVersion")] HRESULT ServerVersion([in] LONGLONG newVal);

    //获取本地一些属性，内部使用
    [id(91), helpstring("method GetLocalFlags")] HRESULT GetLocalFlags([out,retval] LONG* flags);
    [id(92), helpstring("method SetLocalFlags")] HRESULT SetLocalFlags([in] LONG flags);

    //设置服务器的一些属性，内部使用
    [id(93), helpstring("method SetServerDataInfo")] HRESULT SetServerDataInfo([in] DATE tDataModified, [in] BSTR bstrDataMD5);
};
```
