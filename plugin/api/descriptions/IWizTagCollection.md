IWizTagCollection是标签对象IWizTag的一个集合（数据），当需要返回一个IWizTag集合的时候，会返回这个对象，例如IWizDocument.Tags属性，就会返回这个对象。

这个对象实现了IEnumXXXX 接口，这样在某些语言内，可以通过fore_each来枚举集合内的每一个对象。

IWizTagGroupCollection是标签对象IWizTagGroup的一个集合（数据），当需要返回一个IWizTagGroup集合的时候，会返回这个对象，例如IWizDatabase.TagGroups属性，就会返回这个对象。

这个对象实现了IEnumXXXX 接口，这样在某些语言内，可以通过fore_each来枚举集合内的每一个对象。

文件：WizKMCore.dll

```
IWizTagCollection
[
    object,
    uuid(9B636DAC-CCF9-481E-8519-5662F35C6B0D),
    dual,
    nonextensible,
    helpstring("IWizTagCollection Interface"),
    pointer_default(unique)
]
interface IWizTagCollection : IDispatch{

    //生成一个新的IEnumXXXX类型的接口，可以在某些语言内使用for_each类型的语法。
    [id(DISPID_NEWENUM), propget] HRESULT _NewEnum([out, retval] IUnknown** ppUnk);


    //获得某一个对象。Index：索引值，以0开始；返回值：IWizTag对象
    [id(DISPID_VALUE), propget] HRESULT Item([in] long Index, [out, retval] IDispatch** pVal);

    //获得集合内元素的数量
    [id(1), propget] HRESULT Count([out, retval] long * pVal);

    //添加一个标签
    [id(2), helpstring("method Add")] HRESULT Add([in] IDispatch* pTagDisp);
};
```
