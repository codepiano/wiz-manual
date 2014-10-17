IWizTag是WizKMCore.dll包含的COM对象。标签对象必须隶属于一个数据库，因此您不能直接创建标签对象，而是需要通过IWizDatabase对象来获得数据库中的标签对象。

通过IWizDatabase.Tags，可以获得数据库的标签信息，通过IWizDatabase.TagGroups，可以获得数据库中的标签组信息。

文件：WizKMCore.dll

```
[
    object,
    uuid(C710180C-D3F6-4F25-B981-B6820B01C67C),
    dual,
    nonextensible,
    helpstring("IWizTag Interface"),
    pointer_default(unique)
]
interface IWizTag : IDispatch{

    //获得标签的GUID
    [propget, id(1), helpstring("property GUID")] HRESULT GUID([out, retval] BSTR* pVal);

    //获得/设置标签的父标签，类型为IWizTag，如果是跟标签，则返回NULL。
    [propget, id(2), helpstring("property ParentTag")] HRESULT ParentTag([out, retval] IDispatch** pVal);
    [propput, id(2), helpstring("property ParentTag")] HRESULT ParentTag([in ] IDispatch* newVal);

    //获得/设置标签的名称
    [propget, id(3), helpstring("property Name")] HRESULT Name([out, retval] BSTR* pVal);
    [propput, id(3), helpstring("property Name")] HRESULT Name([in] BSTR newVal);

    //获得/设置标签的描述
    [propget, id(4), helpstring("property Description")] HRESULT Description([out, retval] BSTR* pVal);
    [propput, id(4), helpstring("property Description")] HRESULT Description([in] BSTR newVal);

    //获得标签的须改时间
    [propget, id(5), helpstring("property DateModified")] HRESULT DateModified([out, retval] DATE* pVal);

    //获得具有该标签的所有文档
    [propget, id(6), helpstring("property Documents")] HRESULT Documents([out, retval] IDispatch** pVal);

    //获得标签所属的标签组的GUID
    [propget, id(7), helpstring("property TagGroupGUID")] HRESULT TagGroupGUID([out, retval] BSTR* pVal);

    //获得/设置标签的版本。保留
    [propget, id(8), helpstring("property Version")] HRESULT Version([out, retval] LONGLONG* pVal);
    [propput, id(8), helpstring("property Version")] HRESULT Version([in] LONGLONG newVal);

    //获得标签所属的数据库
    [propget, id(9), helpstring("property Database")] HRESULT Database([out, retval] IDispatch** pVal);

    //删除标签
    [id(10), helpstring("method Delete")] HRESULT Delete(void);

    //重新从数据库中获得标签的信息
    [id(11), helpstring("method Reload")] HRESULT Reload(void);

    //获得子标签Collection 返回的是IWizTagCollection
    [propget, id(12), helpstring("property Children")] HRESULT Children([out, retval] IDispatch** pVal);

    //创建子标签，参数分别是名称，描述，返回IWizTag
    [id(13), helpstring("method CreateChildTag")] HRESULT CreateChildTag([in] BSTR bstrTagName, [in] BSTR bstrTagDescription, [out,retval] IDispatch** ppNewTagDisp);

    //移动到根目录，将当前标签作为一个跟标签
    [id(14), helpstring("method MoveToRoot")] HRESULT MoveToRoot();

    //判断一个标签是否在另外一个标签下面，第一个参数：IWizTag，返回true或者false
    [id(15), helpstring("method IsIn")] HRESULT IsIn([in] IDispatch* pFolderDisp, [out,retval] VARIANT_BOOL* pvbRet);

    //获得标签的显示名
    [propget, id(16), helpstring("property DisplayName")] HRESULT DisplayName([out, retval] BSTR* pVal);

    //获得标签的父标签的GUID
    [propget, id(17), helpstring("property ParentGUID")] HRESULT ParentGUID([out, retval] BSTR* pVal);

    //获得标签的完整路径，例如 /Tag1/Tag2/Tag3/
    [propget, id(18), helpstring("property FullPath")] HRESULT FullPath([out, retval] BSTR* pVal);

    //通过标签路径获得对应得标签，可以选择如果对应的标签不存在，是否自动创建
    [id(19), helpstring("method GetChildTagByPath")] HRESULT GetChildTagByPath([in] BSTR bstrTagPath, [in] VARIANT_BOOL vbCreate, [out,retval] IDispatch** ppNewTagDisp);

    //获得标签的拼音
    [id(20), helpstring("method GetPinYin")] HRESULT GetPinYin([in] LONG flags, [out,retval] BSTR* pbstrRetPinYin);

    //获得一个标签下面所有的字标签，包含子标签的字标签
    [propget, id(21), helpstring("property AllChildren")] HRESULT AllChildren([out, retval] IDispatch** pVal);

};
```
