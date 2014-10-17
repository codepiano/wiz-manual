IWizDocumentListCtrl是WizKMControls.dll所包含的一个ActiveX控件，利用这个控件，可以显示Wiz的文档列表，以便用户选择一个或者多个Wiz文档。同时，这个控件还包含了各种用户操作，用户通过右键菜单，可以实现多种操作。

您可以在网页里面直接使用这个控件，也可以在其它的高级语言里面使用，例如C++，VB，C#，Delphi等等。

ProgID：WizKMControls.WizDocumentListCtrl

文件：WizKMControls.dll

```
[
    object,
    uuid(C128ECE0-A006-4E57-8054-4CBC49818231),
    dual,
    nonextensible,
    helpstring("IWizDocumentListCtrl Interface"),
    pointer_default(unique)
]
interface IWizDocumentListCtrl : IDispatch{
    //事件监听

    [propget, id(1), helpstring("property EventsListener")] HRESULT EventsListener([out, retval] VARIANT* pVal);
    [propput, id(1), helpstring("property EventsListener")] HRESULT EventsListener([in] VARIANT newVal);

    //主程序接口
    [propget, id(2), helpstring("property App")] HRESULT App([out, retval] IDispatch** pVal);
    [propput, id(2), helpstring("property App")] HRESULT App([in] IDispatch* newVal);

    //设置/获取数据库对象
    [propget, id(3), helpstring("property Database")] HRESULT Database([out, retval] IDispatch** pVal);
    [propput, id(3), helpstring("property Database")] HRESULT Database([in] IDispatch* newVal);

    //设置/获得选中的笔记
    [propget, id(4), helpstring("property SelectedDocuments")] HRESULT SelectedDocuments([out, retval] IDispatch** pVal);
    [propput, id(4), helpstring("property SelectedDocuments")] HRESULT SelectedDocuments([in] IDispatch* newVal);

    //获得/设置状态section
    [propget, id(5), helpstring("property StateSection")] HRESULT StateSection([out, retval] BSTR* pVal);
    [propput, id(5), helpstring("property StateSection")] HRESULT StateSection([in] BSTR newVal);

    //是否显示边框，已经失效
    [propget, id(6), helpstring("property Border")] HRESULT Border([out, retval] VARIANT_BOOL* pVal);
    [propput, id(6), helpstring("property Border")] HRESULT Border([in] VARIANT_BOOL newVal);

    //是否禁止右键菜单
    [propget, id(7), helpstring("property DisableContextMenu")] HRESULT DisableContextMenu([out, retval] VARIANT_BOOL* pVal);
    [propput, id(7), helpstring("property DisableContextMenu")] HRESULT DisableContextMenu([in] VARIANT_BOOL newVal);


    //获得/设置排序依据，例如Title，URL等等
    [propget, id(8), helpstring("property SortBy")] HRESULT SortBy([out, retval] BSTR* pVal);
    [propput, id(8), helpstring("property SortBy")] HRESULT SortBy([in] BSTR newVal);

    //获得/设置排序顺序，从大到小还是从小到大
    [propget, id(9), helpstring("property SortOrder")] HRESULT SortOrder([out, retval] BSTR* pVal);
    [propput, id(9), helpstring("property SortOrder")] HRESULT SortOrder([in] BSTR newVal);

    //获得文档列表所属的父文件夹，如果是搜索列表，则为空。类型为IWizFolder
    [propget, id(10), helpstring("property ParentFolder")] HRESULT ParentFolder([out, retval] IDispatch** pVal);

    //设置文档列表，pDisp：IWizDocumentCollection
    [id(11), helpstring("method SetDocuments")] HRESULT SetDocuments([in] IDispatch* pDisp);

    //设置文档列表。pDisp：IWizDocumentCollection对象；bstrSortBy：排序依据；bstrSortOrder：排序方式
    [id(12), helpstring("method SetDocuments2")] HRESULT SetDocuments2([in] IDispatch* pDisp, [in] BSTR bstrSortBy, [in] BSTR bstrSortOrder);

    //刷新
    [id(13), helpstring("method Refresh")] HRESULT Refresh();

    //获得所有文档
    [id(14), helpstring("method GetDocuments")] HRESULT GetDocuments([out, retval] IDispatch** pVal);

    //获得/设置用户选中的文档更改的事件，提供给html+JavaScript使用这个事件
    [propget, id(15), helpstring("property OnSelChanged")] HRESULT OnSelChanged([out, retval] VARIANT* pVal);
    [propput, id(15), helpstring("property OnSelChanged")] HRESULT OnSelChanged([in] VARIANT newVal);

    //获得/设置用户双击一个文档的事件，提供给html+JavaScript使用这个事件
    [propget, id(16), helpstring("property OnDblClickItem")] HRESULT OnDblClickItem([out, retval] VARIANT* pVal);
    [propput, id(16), helpstring("property OnDblClickItem")] HRESULT OnDblClickItem([in] VARIANT newVal);

    //获得/设置用户选中一个右键菜单的事件，提供给html+JavaScript使用这个事件
    [propget, id(17), helpstring("property OnExecCommand")] HRESULT OnExecCommand([out, retval] VARIANT* pVal);
    [propput, id(17), helpstring("property OnExecCommand")] HRESULT OnExecCommand([in] VARIANT newVal);


    //设置/获取一些属性
    [propget, id(15), helpstring("property Options")] HRESULT Options([out, retval] LONG* pVal);
    [propput, id(15), helpstring("property Options")] HRESULT Options([in] LONG newVal);

    //第二行类型
    [propget, id(16), helpstring("property SecondLineType")] HRESULT SecondLineType([out, retval] LONG* pVal);
    [propput, id(16), helpstring("property SecondLineType")] HRESULT SecondLineType([in] LONG newVal);

    //设置过滤器
    [id(17), helpstring("method SetFilter")] HRESULT SetFilter([in] IUnknown* pFilterUnk);

    //刷新笔记
    [id(18), helpstring("method RefreshDocuments")] HRESULT RefreshDocuments();

    //设置/获得显示类型
    [propget, id(19), helpstring("property ViewType")] HRESULT ViewType([out, retval] LONG* pVal);
    [propput, id(19), helpstring("property ViewType")] HRESULT ViewType([in] LONG newVal);

}
```
