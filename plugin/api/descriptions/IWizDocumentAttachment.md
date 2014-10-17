IWizDocumentAttachment对象，文档的附件。

文件：WizKMCore.dll

```
[
    object,
    uuid(9368D723-8B76-45C3-9B65-BD88B0A1BE64),
    dual,
    nonextensible,
    helpstring("IWizDocumentAttachment Interface"),
    pointer_default(unique)
]
interface IWizDocumentAttachment : IDispatch{

    //获得附件的GUID
    [propget, id(1), helpstring("property GUID")] HRESULT GUID([out, retval] BSTR* pVal);

    //获得/设置附件的名称
    [propget, id(2), helpstring("property Name")] HRESULT Name([out, retval] BSTR* pVal);
    [propput, id(2), helpstring("property Name")] HRESULT Name([in] BSTR newVal);

    //获得/设置附件的描述
    [propget, id(3), helpstring("property Description")] HRESULT Description([out, retval] BSTR* pVal);
    [propput, id(3), helpstring("property Description")] HRESULT Description([in] BSTR newVal);

    //获得/设置附件的原始URL
    [propget, id(4), helpstring("property URL")] HRESULT URL([out, retval] BSTR* pVal);
    [propput, id(4), helpstring("property URL")] HRESULT URL([in] BSTR newVal);

    //获得/设置附件的文件大小，单位是字节。
    [propget, id(5), helpstring("property Size")] HRESULT Size([out, retval] LONG* pVal);

    //获得附件的信息修改日期
    [propget, id(6), helpstring("property InfoDateModified")] HRESULT InfoDateModified([out, retval] DATE* pVal);

    //获得附件的信息md5值
    [propget, id(7), helpstring("property InfoMD5")] HRESULT InfoMD5([out, retval] BSTR* pVal);

    //获得附件的数据修改日期
    [propget, id(8), helpstring("property DataDateModified")] HRESULT DataDateModified([out, retval] DATE* pVal);

    //获得附件的数据md5值
    [propget, id(9), helpstring("property DataMD5")] HRESULT DataMD5([out, retval] BSTR* pVal);

    //获得附件所属的文档
    [propget, id(10), helpstring("property Document")] HRESULT Document([out, retval] IDispatch** pVal);

    //获得附件所属的文档GUID
    [propget, id(11), helpstring("property DocumentGUID")] HRESULT DocumentGUID([out, retval] BSTR* pVal);

    //获得附件的磁盘文件名
    [propget, id(12), helpstring("property FileName")] HRESULT FileName([out, retval] BSTR* pVal);

    //获得附件的版本，用于同步，保留
    [propget, id(13), helpstring("property Version")] HRESULT Version([out, retval] LONGLONG* pVal);
    [propput, id(13), helpstring("property Version")] HRESULT Version([in] LONGLONG newVal);

    //删除附件
    [id(14), helpstring("method Delete")] HRESULT Delete(void);

    //重新从数据库中获得附件信息
    [id(15), helpstring("method Reload")] HRESULT Reload(void);

    //更新附件的数据md5值。
    [id(16), helpstring("method UpdateDataMD5")] HRESULT UpdateDataMD5(void);

    //获得/设置附件数据已经是否已经被下载到本地

    [propget, id(17), helpstring("property Downloaded")] HRESULT Downloaded([out, retval] VARIANT_BOOL* pVal);
    [propput, id(17), helpstring("property Downloaded")] HRESULT Downloaded([in] VARIANT_BOOL newVal);

    //检查附件数据是否被下载，如果没有，则自动下载。
    [id(18), helpstring("method CheckAttachmentData")] HRESULT CheckAttachmentData([in] LONGLONG hwndParent, [out,retval] VARIANT_BOOL* vbRet);

    //获取/设置附件数据，IStream类型
    [id(19), helpstring("method SetData")] HRESULT SetData([in] IUnknown* pData);
    [id(20), helpstring("method GetData")] HRESULT GetData([in] IUnknown* pData);


    //获取/设置笔记的服务器版本号，内部使用
    [propget, id(21), helpstring("property ServerVersion")] HRESULT ServerVersion([out, retval] LONGLONG* pVal);
    [propput, id(21), helpstring("property ServerVersion")] HRESULT ServerVersion([in] LONGLONG newVal);

    //获取本地一些属性，内部使用
    [id(22), helpstring("method GetLocalFlags")] HRESULT GetLocalFlags([out,retval] LONG* flags);
    [id(23), helpstring("method SetLocalFlags")] HRESULT SetLocalFlags([in] LONG flags);

    //设置服务器的一些属性，内部使用
    [id(24), helpstring("method SetServerDataInfo")] HRESULT SetServerDataInfo([in] DATE tDataModified, [in] BSTR bstrDataMD5);

};
```
