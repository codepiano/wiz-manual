IWizDocumentParam对象。每一个文档，都可以包含任意条自定义参数。自定义参数，可以用来完成各种扩展功能，尤其是进行二次开发的时候，会比较多的用到文档参数。

文档参数也会和服务器同步。

文件：WizKMCore.dll

```
[
    object,
    uuid(46C5786E-A72E-4089-B4DB-70B3DD174FEB),
    dual,
    nonextensible,
    helpstring("IWizDocumentParam Interface"),
    pointer_default(unique)
]
interface IWizDocumentParam : IDispatch{

    //文档参数名称
    [propget, id(1), helpstring("property Name")] HRESULT Name([out, retval] BSTR* pVal);

    //文档参数的值
    [propget, id(2), helpstring("property Value")] HRESULT Value([out, retval] BSTR* pVal);

    //获得param所附属的文档
    [propget, id(3), helpstring("property Document")] HRESULT Document([out, retval] IDispatch** pVal);

    //删除这个参数
    [id(4), helpstring("method Delete")] HRESULT Delete(void);
};
```
