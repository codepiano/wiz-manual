IWizStyle是WizKMCore.dll包含的一个COM对象。样式对象必须隶属于一个数据库，因此您不能直接创建这个对象，而是需要通过IWizDatabase对象来获得数据库中的样式对象。

通过IWizDatabase.Styles，可以获得数据库的样式信息。

一个文档只能有一个样式，或者没有。样式主要用来突出显示某些特殊的文档，可以用来给文档进行分类。当点击某一个样式的时候，会列出所有具有这个样式的文档。

文件：WizKMCore.dll

```
[
    object,
    uuid(77C6F00C-6A41-4A14-B3E6-986401D5A5D5),
    dual,
    nonextensible,
    helpstring("IWizStyle Interface"),
    pointer_default(unique)
]
interface IWizStyle : IDispatch{

    //获得样式的GUID
    [propget, id(1), helpstring("property GUID")] HRESULT GUID([out, retval] BSTR* pVal);

    //获得/设置样式的名称
    [propget, id(2), helpstring("property Name")] HRESULT Name([out, retval] BSTR* pVal);
    [propput, id(2), helpstring("property Name")] HRESULT Name([in] BSTR newVal);

    //获得/设置样式的样式
    [propget, id(3), helpstring("property Description")] HRESULT Description([out, retval] BSTR* pVal);
    [propput, id(3), helpstring("property Description")] HRESULT Description([in] BSTR newVal);

    //获得/设置样式的文字颜色
    [propget, id(4), helpstring("property TextColor")] HRESULT TextColor([out, retval] LONG* pVal);
    [propput, id(4), helpstring("property TextColor")] HRESULT TextColor([in] LONG newVal);

    //获得/设置样式的背景色
    [propget, id(5), helpstring("property BackColor")] HRESULT BackColor([out, retval] LONG* pVal);
    [propput, id(5), helpstring("property BackColor")] HRESULT BackColor([in] LONG newVal);


    //获得/设置样式的文字是否粗体显示
    [propget, id(6), helpstring("property TextBold")] HRESULT TextBold([out, retval] VARIANT_BOOL* pVal);
    [propput, id(6), helpstring("property TextBold")] HRESULT TextBold([in] VARIANT_BOOL newVal);

    //获得/设置样式的旗帜索引，暂不支持
    [propget, id(7), helpstring("property FlagIndex")] HRESULT FlagIndex([out, retval] LONG* pVal);
    [propput, id(7), helpstring("property FlagIndex")] HRESULT FlagIndex([in] LONG newVal);

    //获得样式的修改时间
    [propget, id(8), helpstring("property DateModified")] HRESULT DateModified([out, retval] DATE* pVal);

    //获得所有具有这个样式的文档，类型为IWizDocumentCollection
    [propget, id(9), helpstring("property Documents")] HRESULT Documents([out, retval] IDispatch** pVal);

    //获得/设置样式的版本，保留
    [propget, id(10), helpstring("property Version")] HRESULT Version([out, retval] LONGLONG* pVal);
    [propput, id(10), helpstring("property Version")] HRESULT Version([in] LONGLONG newVal);

    //获得样式所属的数据库
    [propget, id(11), helpstring("property Database")] HRESULT Database([out, retval] IDispatch** pVal);

    //删除样式
    [id(12), helpstring("method Delete")] HRESULT Delete(void);

    //从数据库中重新获得样式的信息
    [id(13), helpstring("method Reload")] HRESULT Reload(void);
};
```
