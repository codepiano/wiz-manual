IWizFolder是WizKMCore.dll包含的一个COM对象。文件夹对象必须隶属于一个数据库，因此您不能直接创建这个对象，而是需要通过IWizDatabase对象来获得数据库中的文件夹对象。

通过IWizDatabase.Folders，可以获得数据库的所有根文件夹，进而获得其它的子文件夹信息。

每一个IWizFolder对象，都对应于用户数据库文件夹下面的一个子文件夹，如下图：

![用户目录](../img/user-folder.jpg)

用户可以自行在数据库文件夹内建立新的文件夹，或者删除空的文件夹，Wiz会自动显示。但是请不要删除包含文档的文件夹，否则会引起数据不一致，因为Wiz对文档进行了索引。

文件：WizKMCore.dll

```
[
    object,
    uuid(B0E49F7C-CF4B-4DD4-A3F3-432F628C3A31),
    dual,
    nonextensible,
    helpstring("IWizFolder Interface"),
    pointer_default(unique)
]
interface IWizFolder : IDispatch{

    //获得或者设置文件夹名称，直接对应于磁盘文件夹的名称。
    [propget, id(1), helpstring("property Name")] HRESULT Name([out, retval] BSTR* pVal);
    [propput, id(1), helpstring("property Name")] HRESULT Name([in] BSTR newVal);

    //获得或者设置文件夹是否需要和网络同步。仅针对根文件夹有效。
    [propget, id(2), helpstring("property Sync")] HRESULT Sync([out, retval] VARIANT_BOOL* pVal);
    [propput, id(2), helpstring("property Sync")] HRESULT Sync([in] VARIANT_BOOL newVal);

    //获得完整的文件夹路径，例如D:\My Documents\My Knowledge\Data\Default\新闻
    [propget, id(3), helpstring("property FullPath")] HRESULT FullPath([out, retval] BSTR* pVal);

    //获得文件夹的位置，相对于数据库文件夹，格式为/abc/def/类型
    [propget, id(4), helpstring("property Location")] HRESULT Location([out, retval] BSTR* pVal);

    //获得文件夹内所有文档的集合，类型为IWizDocumentCollection
    [propget, id(5), helpstring("property Documents")] HRESULT Documents([out, retval] IDispatch** pVal);

    //获得所有子文件夹的集合。类型为IWizFolderCollection
    [propget, id(6), helpstring("property Folders")] HRESULT Folders([out, retval] IDispatch** pVal);

    //是否是根文件夹
    [propget, id(7), helpstring("property IsRoot")] HRESULT IsRoot([out, retval] VARIANT_BOOL* pVal);

    //是否是已删除文件夹（回收站）
    [propget, id(8), helpstring("property IsDeletedItems")] HRESULT IsDeletedItems([out, retval] VARIANT_BOOL* pVal);

    //获得父文件夹，如果是个文件夹，返回值为空。返回类型为IWizFolder
    [propget, id(9), helpstring("property Parent")] HRESULT Parent([out, retval] IDispatch** pVal);

    //获得文件夹所在的数据库对象，类型为IWizDatabase
    [propget, id(10), helpstring("property Database")] HRESULT Database([out, retval] IDispatch** pVal);

    //创建子文件夹。bstrFolderName：子文件夹名称；返回值：成功创建的文件夹对象，类型为IWizFolder
    [id(11), helpstring("method CreateSubFolder")] HRESULT CreateSubFolder([in] BSTR bstrFolderName, [out,retval] IDispatch** ppNewFolderDisp);

    //创建一个新的文档。bstrTitle：文档标题；strName：文档名称，一般就是文档对应的磁盘ziw文件名，如果没有指定，才用标题作为文件名；bstrURL：文档的原始URL；返回值：成功创建的文档对象，类型为IWizDocument
    [id(12), helpstring("method CreateDocument")] HRESULT CreateDocument([in] BSTR bstrTitle, [in] BSTR strName, [in] BSTR bstrURL, [out,retval] IDispatch** ppNewDocumentDisp);
    [id(13), helpstring("method CreateDocument2")] HRESULT CreateDocument2([in] BSTR bstrTitle, [in] BSTR bstrURL, [out,retval] IDispatch** ppNewDocumentDisp);

    //删除文件夹
    [id(14), helpstring("method Delete")] HRESULT Delete(void);

    //移动到其它的文件夹。pDestFolderDisp：目标文件夹
    [id(15), helpstring("method MoveTo")] HRESULT MoveTo([in] IDispatch* pDestFolderDisp);

    //移动到跟目录，当前文件夹作为数据库下面的根文件夹
    [id(16), helpstring("method MoveToRoot")] HRESULT MoveToRoot();

    //是否是目标文件夹的子文件夹（多级）。pFolderDisp：目标文件夹
    [id(17), helpstring("method IsIn")] HRESULT IsIn([in] IDispatch* pFolderDisp, [out,retval] VARIANT_BOOL* pvbRet);

    //设置文件夹自定义图标。hIcon：图标句柄，类型为HICON。该句柄仅针对进城内调用有效
    [id(18), helpstring("method SetIcon")] HRESULT SetIcon([in] OLE_HANDLE hIcon);

    //设置文件夹自定义图标。bstrIconFileName：图标文件名，可以是icon、exe、dll等可以包含图标的文件；nIconIndex：图标索引，对于exe、dll文件，可能包含多个图标，可以指定是哪一个图标。如果nIconIndex大于等于0，表示图标索引，如果小于0，表示是图标的资源ID。
    [id(19), helpstring("method SetIcon2")] HRESULT SetIcon2([in] BSTR bstrIconFileName, [in] LONG nIconIndex);

    //获得文件夹图标文件名。
    [id(20), helpstring("method GetIconFileName")] HRESULT GetIconFileName([out,retval] BSTR* pbstrIconFileName);

    //获得文档数量。nFlags：选项，可以为wizDocumentCountContainsSubFolders=0x02，表示包含子文件夹。
    [id(21), helpstring("method GetDocumentCount")] HRESULT GetDocumentCount([in] LONG nFlags, [out,retval] LONG* pnCount);

    //获得文件夹的显示名称，例如My Journals文件夹，在不同的语言下面，显示名称不同。nFlags：保留，必须为0
    [id(22), helpstring("method GetDisplayName")] HRESULT GetDisplayName([in] LONG nFlags, [out,retval] BSTR* pbstrDisplayName);

    //获得用于显示文件夹的内文档的显示模板。
    [id(23), helpstring("method GetDisplayTemplate")] HRESULT GetDisplayTemplate([out,retval] BSTR* pbstrReaingTemplateFileName);

    //设置用于显示文件夹内的文档的显示模板
    [id(24), helpstring("method SetDisplayTemplate")] HRESULT SetDisplayTemplate([in] BSTR bstrReaingTemplateFileName);

    //获得文件夹所在的根文件夹的对象。返回值类型为IWizFolder。
    [propget, id(25), helpstring("property RootFolder")] HRESULT RootFolder([out, retval] IDispatch** pVal);

    //获得文件家的参数。
    [propget, id(26), helpstring("property Settings")] HRESULT Settings([in] BSTR bstrSection, [in] BSTR bstrKeyName, [out, retval] BSTR* pVal);
    [propput, id(26), helpstring("property Settings")] HRESULT Settings([in] BSTR bstrSection, [in] BSTR bstrKeyName, [in] BSTR newVal);

    //获得文件夹是否被用户设置为加密
    [propget, id(27), helpstring("property Encrypt")] HRESULT Encrypt([out, retval] VARIANT_BOOL* pVal);
    [propput, id(27), helpstring("property Encrypt")] HRESULT Encrypt([in] VARIANT_BOOL newVal);

};
```
