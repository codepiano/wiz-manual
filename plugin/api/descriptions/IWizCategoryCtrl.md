IWizCategoryCtrl是WizKMControls.dll所包含的一个ActiveX控件，利用这个控件，可以显示Wiz的目录，标签，样式，群组以及快速搜索，以便用户选择Wiz文件夹，标签等等。同时，这个控件还包含了各种用户操作，用户通过右键菜单，可以实现多种操作。

您可以在网页里面直接使用这个控件，也可以在其它的高级语言里面使用，例如C++，VB，C#，Delphi等等。

ProgID：WizKMControls.WizCategoryCtrl
文件：WizKMControls.dll

```
[
    object,
    uuid(C163AE7B-E5B1-4789-A61B-27AFC8C9E17D),
    dual,
    nonextensible,
    helpstring("IWizCategoryCtrl Interface"),
    pointer_default(unique)
]
interface IWizCategoryCtrl : IDispatch{
    //事件监听
    [propget, id(1), helpstring("property EventsListener")] HRESULT EventsListener([out, retval] VARIANT* pVal);
    [propput, id(1), helpstring("property EventsListener")] HRESULT EventsListener([in] VARIANT newVal);

    //主程序接口
    [propget, id(2), helpstring("property App")] HRESULT App([out, retval] IDispatch** pVal);
    [propput, id(2), helpstring("property App")] HRESULT App([in] IDispatch* newVal);

    //数据库
    [propget, id(3), helpstring("property Database")] HRESULT Database([out, retval] IDispatch** pVal);
    [propput, id(3), helpstring("property Database")] HRESULT Database([in] IDispatch* newVal);

    //选中的类型
    [propget, id(4), helpstring("property SelectedType")] HRESULT SelectedType([out, retval] LONG* pVal);

    //选中的文件夹
    [propget, id(5), helpstring("property SelectedFolder")] HRESULT SelectedFolder([out, retval] IDispatch** pVal);
    [propput, id(5), helpstring("property SelectedFolder")] HRESULT SelectedFolder([in] IDispatch* newVal);

    //选中的标签
    [propget, id(6), helpstring("property SelectedTags")] HRESULT SelectedTags([out, retval] IDispatch** pVal);
    [propput, id(6), helpstring("property SelectedTags")] HRESULT SelectedTags([in] IDispatch* newVal);

    //选中的样式
    [propget, id(7), helpstring("property SelectedStyle")] HRESULT SelectedStyle([out, retval] IDispatch** pVal);
    [propput, id(7), helpstring("property SelectedStyle")] HRESULT SelectedStyle([in] IDispatch* newVal);

    //选中的笔记
    [propget, id(8), helpstring("property SelectedDocument")] HRESULT SelectedDocument([out, retval] IDispatch** pVal);
    [propput, id(8), helpstring("property SelectedDocument")] HRESULT SelectedDocument([in] IDispatch* newVal);

    //设置section
    [propget, id(9), helpstring("property StateSection")] HRESULT StateSection([out, retval] BSTR* pVal);
    [propput, id(9), helpstring("property StateSection")] HRESULT StateSection([in] BSTR newVal);

    //属性
    [propget, id(10), helpstring("property Options")] HRESULT Options([out, retval] LONG* pVal);
    [propput, id(10), helpstring("property Options")] HRESULT Options([in] LONG newVal);

    //设置/获得是否显示笔记
    [propget, id(11), helpstring("property ShowDocuments")] HRESULT ShowDocuments([out, retval] VARIANT_BOOL* pVal);
    [propput, id(11), helpstring("property ShowDocuments")] HRESULT ShowDocuments([in] VARIANT_BOOL newVal);

    //设置是否显示边框，已经失效
    [propget, id(12), helpstring("property Border")] HRESULT Border([out, retval] VARIANT_BOOL* pVal);
    [propput, id(12), helpstring("property Border")] HRESULT Border([in] VARIANT_BOOL newVal);

    //刷新
    [id(13), helpstring("method Refresh")] HRESULT Refresh([in] LONG nFlags);

    //执行命令
    [id(14), helpstring("method ExecuteCommand")] HRESULT ExecuteCommand([in] BSTR bstrCommandName, [in] VARIANT* pvParam1, [in] VARIANT* pvParam2, [out, retval] VARIANT* pvRetValue);

    //获得选中的笔记
    [id(15), helpstring("method GetSelectedDocuments2")] HRESULT GetSelectedDocuments2([out] BSTR* pbstrSortBy, [out] BSTR* pbstrSortOrder, [out, retval] IDispatch** pVal);

    //保存状态
    [id(16), helpstring("method SaveState")] HRESULT SaveState();

    //已经失效
    [id(17), helpstring("method GetSelectedItemCustomData")] HRESULT GetSelectedItemCustomData([out, retval] BSTR* pbstrData);

    //获得当前数据库
    [propget, id(18), helpstring("property CurrentDatabase")] HRESULT CurrentDatabase([out, retval] IDispatch** pVal);

    //设置搜索结果
    [id(19), helpstring("method SetSearchResult")] HRESULT SetSearchResult([in] IDispatch* pDatabaseDisp, [in] BSTR bstrKeywords, [in] IDispatch* pDocumentDisp, [in] VARIANT_BOOL vbForceSelect);
};
```
