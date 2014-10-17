IWizDocumenCollection是文件夹对象IWizDocumen的一个集合（数据），当需要返回一个IWizDocumen集合的时候，会返回这个对象，例如IWizFolder.Documents属性，就会返回这个对象。
这个对象实现了IEnumXXXX 接口，这样在某些语言内，可以通过fore_each来枚举集合内的每一个对象。

文件：WizKMCore.dll

```
[
    object,
    uuid(F8AE7F18-3C30-47D1-8902-947E5676895C),
    dual,
    nonextensible,
    helpstring("IWizDocumentCollection Interface"),
    pointer_default(unique)
]
interface IWizDocumentCollection : IDispatch{

    //生成一个新的IEnumXXXX类型的接口，可以在某些语言内使用for_each类型的语法。
    [id(DISPID_NEWENUM), propget] HRESULT _NewEnum([out, retval] IUnknown** ppUnk);

    //获得某一个对象。Index：索引值，以0开始；返回值：IWizDocument对象
    [id(DISPID_VALUE), propget] HRESULT Item([in] long Index, [out, retval] IDispatch** pVal);

    //获得集合内元素的数量
    [id(1), propget] HRESULT Count([out, retval] long * pVal);

    //添加一个文档，类型必须为IWizDocument
    [id(2), helpstring("method Add")] HRESULT Add([in] IDispatch* pDocumentDisp);

    //删除某一个元素

    [id(3), helpstring("method Delete")] HRESULT Delete([in] LONG nIndex);

    //清空列表
    [id(4), helpstring("method Clear")] HRESULT Clear();

};
```
