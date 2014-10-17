IWizStyleCollection是文件夹对象IWizStyle的一个集合（数据），当需要返回一个IWizStyle集合的时候，会返回这个对象，例如IWizDatabase.Styles属性，就会返回这个对象。

这个对象实现了IEnumXXXX 接口，这样在某些语言内，可以通过fore_each来枚举集合内的每一个对象。

文件：WizKMCore.dll

```
[
    object,
    uuid(E6A5A90D-5766-428F-BE25-AFC9D641CB88),
    dual,
    nonextensible,
    helpstring("IWizStyleCollection Interface"),
    pointer_default(unique)
]
interface IWizStyleCollection : IDispatch{

    //生成一个新的IEnumXXXX类型的接口，可以在某些语言内使用for_each类型的语法。
    [id(DISPID_NEWENUM), propget] HRESULT _NewEnum([out, retval] IUnknown** ppUnk);


    //获得某一个对象。Index：索引值，以0开始；返回值：IWizStyle对象
    [id(DISPID_VALUE), propget] HRESULT Item([in] long Index, [out, retval] IDispatch** pVal);


    //获得集合内元素的数量
    [id(1), propget] HRESULT Count([out, retval] long * pVal);
};
```
