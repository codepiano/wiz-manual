IWizDocumentAttachmentCollection是IWizDocumentAttachment的集合

文件：WizKMCore.dll

```
[
    object,
    uuid(9B1E33B1-C799-4F7F-8926-54A512D45635),
    dual,
    nonextensible,
    helpstring("IWizDocumentAttachmentCollection Interface"),
    pointer_default(unique)
]
interface IWizDocumentAttachmentCollection : IDispatch{

    //生成一个新的IEnumXXXX类型的接口，可以在某些语言内使用for_each类型的语法。
    [id(DISPID_NEWENUM), propget] HRESULT _NewEnum([out, retval] IUnknown** ppUnk);

    //获得某一个对象。Index：索引值，以0开始；返回值：IWizDocumentAttachment对象
    [id(DISPID_VALUE), propget] HRESULT Item([in] long Index, [out, retval] IDispatch** pVal);

    //获得集合内元素的数量
    [id(1), propget] HRESULT Count([out, retval] long * pVal);

    //添加一个附件对象，类型为IWizDocumentAttachment
    [id(2), helpstring("method Add")] HRESULT Add([in] IDispatch* pDocumentAttachmentDisp);
};
```
