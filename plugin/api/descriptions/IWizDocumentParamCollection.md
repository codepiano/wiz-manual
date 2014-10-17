IWizDocumentParamCollection是IWizDocumentParam对象的集合。

文件：WizKMCore.dll

```
[
    object,
    uuid(8DB6192A-1C66-4DD0-ABD4-E80C12969A13),
    dual,
    nonextensible,
    helpstring("IWizDocumentParamCollection Interface"),
    pointer_default(unique)
]
interface IWizDocumentParamCollection : IDispatch{

    //生成一个新的IEnumXXXX类型的接口，可以在某些语言内使用for_each类型的语法。
    [id(DISPID_NEWENUM), propget] HRESULT _NewEnum([out, retval] IUnknown** ppUnk);

    //获得某一个对象。Index：索引值，以0开始；返回值：IWizDocumentParam对象
    [id(DISPID_VALUE), propget] HRESULT Item([in] long Index, [out, retval] IDispatch** pVal);

    //获得集合内元素的数量
    [id(1), propget] HRESULT Count([out, retval] long * pVal);
};
```
