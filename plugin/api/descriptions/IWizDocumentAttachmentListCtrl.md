IWizDocumentAttachmentListCtrl是WizKMControls.dll所包含的一个ActiveX控件，利用这个控件，可以显示Wiz的文档列表，以便用户选择一个或者多个Wiz文档。同时，这个控件还包含了各种用户操作，用户通过右键菜单，可以实现多种操作。

您可以在网页里面直接使用这个控件，也可以在其它的高级语言里面使用，例如C++，VB，C#，Delphi等等。

ProgID：WizKMControls.WizDocumentAttachmentListCtrl

文件：WizKMControls.dll

```
[
    object,
    uuid(39B2717D-7FDA-4EDD-91A4-0173FD35B871),
    dual,
    nonextensible,
    helpstring("IWizDocumentAttachmentListCtrl Interface"),
    pointer_default(unique)
]
interface IWizDocumentAttachmentListCtrl : IDispatch{

    //设置/获取事件监听消息
    [propget, id(1), helpstring("property EventsListener")] HRESULT EventsListener([out, retval] VARIANT* pVal);
    [propput, id(1), helpstring("property EventsListener")] HRESULT EventsListener([in] VARIANT newVal);

    //设置/获取主程序APP
    [propget, id(2), helpstring("property App")] HRESULT App([out, retval] IDispatch** pVal);
    [propput, id(2), helpstring("property App")] HRESULT App([in] IDispatch* newVal);

     //获得/设置数据库对象，类型为IWizDatabase
    [propget, id(3), helpstring("property Database")] HRESULT Database([out, retval] IDispatch** pVal);
    [propput, id(3), helpstring("property Database")] HRESULT Database([in] IDispatch* newVal);

    //获得/设置附件列表所属的文档对象，类型为IWizDocument
    [propget, id(4), helpstring("property Document")] HRESULT Document([out, retval] IDispatch** pVal);
    [propput, id(4), helpstring("property Document")] HRESULT Document([in] IDispatch* newVal);

    //获得/设置用户选中的附件对象，类型为IWizDocumentAttachmentCollection
    [propget, id(5), helpstring("property SelectedAttachments")] HRESULT SelectedAttachments([out, retval] IDispatch** pVal);
    [propput, id(5), helpstring("property SelectedAttachments")] HRESULT SelectedAttachments([in] IDispatch* newVal);

    //获得/设置是否显示边框，已经无效
    [propget, id(6), helpstring("property Border")] HRESULT Border([out, retval] VARIANT_BOOL* pVal);
    [propput, id(6), helpstring("property Border")] HRESULT Border([in] VARIANT_BOOL newVal);

    //执行添加附件命令。
    [id(7), helpstring("method AddAttachments")] HRESULT AddAttachments(void);

    //执行添加附件命令，可以传入一个文件名数组，直接添加某些文件作为附件
    [id(8), helpstring("method AddAttachments2")] HRESULT AddAttachments2([in] VARIANT* pvFiles);

    //或者最小高度
    [id(9), helpstring("method GetMinSize")] HRESULT GetMinSize([out, retval] ULONG* pnsize);
};
```
