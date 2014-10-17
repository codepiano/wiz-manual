该对象提供了访问Wiz主界面的功能，可以获得Wiz的一些界面控件，或者设置一些控件的状态。

idl文件：WizExplorer.idl

```
[
object,
uuid(B15879A3-8647-4d6f-84D8-3A24709EC962),
dual,
nonextensible,
helpstring("IWizExplorerWindow Interface"),
pointer_default(unique)
]
interface IWizExplorerWindow : IDispatch{


//获取左侧分类树对象，IWizCategoryCtrl
[propget, id(1), helpstring("property CategoryCtrl")] HRESULT CategoryCtrl([out, retval] IDispatch** pVal);

//获取中间笔记列表控件对象
[propget, id(2), helpstring("property DocumentsCtrl")] HRESULT DocumentsCtrl([out, retval] IDispatch** pVal);

//获取主窗口窗口句柄
[propget, id(3), helpstring("property HWND")] HRESULT HWND([out, retval] OLE_HANDLE* pVal);

//获得当前Wiz正在显示的Wiz文档对象，可能为null。类型为：WizDocument
[propget, id(4), helpstring("property CurrentDocument")] HRESULT CurrentDocument([out, retval] IDispatch** pVal);

//获得当前Wiz正在显示的html的document对象。在windows下面就是IHTMLDocument2对象。可能为null
[propget, id(5), helpstring("property CurrentDocumentHtmlDocument")] HRESULT CurrentDocumentHtmlDocument([out, retval] IDispatch** pVal);

//获得当前Wiz正在显示的文档对应得附件列表对象。可能为null。类型为：WizDocumentAttachmentListCtrl。
[propget, id(6), helpstring("property CurrentDocumentAttachmentCtrl")] HRESULT CurrentDocumentAttachmentCtrl([out, retval] IDispatch** pVal);

//显示一个html对话框，可以设定标题，url，宽度，高度，扩展参数（必须为""）。其中vRet是CloseHtmlDialog指定的参数
[id(7), helpstring("method ShowHtmlDialog")] HRESULT ShowHtmlDialog([in] BSTR bstrTitle, [in] BSTR bstrURL, [in] LONG nWidth, [in] LONG nHeight, [in] BSTR bstrExtOptions, [out, retval] VARIANT* pnRet);

//关闭当前Html对话框，其中vRet是关闭的结果，为自定义参数，vRet将会作为ShowHtmlDialog返回
[id(8), helpstring("method CloseHtmlDialog")] HRESULT CloseHtmlDialog([in] IDispatch* pHtmlDocumentDisp, [in] VARIANT vRet);

//获得当前HTML对话框的参数，该参数通过ShowHtmlDialog传入
[id(9), helpstring("method GetHtmlDialogParam")] HRESULT GetHtmlDialogParam([in] IDispatch* pHtmlDocumentDisp, [out, retval] VARIANT* pvParam);

//在wiz内部显示一个文档
[id(12), helpstring("method ViewDocument")] HRESULT ViewDocument([in] IDispatch* pDocumentDisp, [in] VARIANT_BOOL vbOpenInNewTab);

//显示一个html文件（不同于html对话框，在Wiz的文档窗口tab里面显示）
[id(13), helpstring("method ViewHtml")] HRESULT ViewHtml([in] BSTR bstrURL, [in] VARIANT_BOOL vbOpenInNewTab);

//获得html对话框的窗口句柄
[id(14), helpstring("method GetHtmlDialogHWND")] HRESULT GetHtmlDialogHWND([in] IDispatch* pHtmlDocumentDisp, [out,retval] OLE_HANDLE* phHWND);

//显示一个消息框。参数和win 32 api的MessageBox参数相同。主要是为了弥补单纯的js文件没有alert/confirm等方法的缺陷。
[id(15), helpstring("method ShowMessage")] HRESULT ShowMessage([in] BSTR bstrText, [in] BSTR bstrTitle, [in] LONG nType, [out,retval] LONG* pnRet);

//显示一个下拉窗口。
[id(16), helpstring("method ShowSelectorWindow")] HRESULT ShowSelectorWindow([in] BSTR bstrURL, [in] LONG left, [in] LONG top, [in] LONG width, [in] LONG height, [in] BSTR bstrExtOptions);

//关闭下拉窗口
[id(17), helpstring("method CloseSelectorWindow")] HRESULT CloseSelectorWindow([in] IDispatch* pdispSelectorHtmlDocument);

//添加一个工具栏按钮，目前仅能添加到文档阅读/编辑窗口上面，在阅读/附件按钮后面。
//bstrType必须是"document"；
//bstrButtonID：按钮ID；
//bstrButtonText：按钮文字；
//bstrIconFileName：按钮上面图标的文件，可以指定一个ico文件（包含路径）；
//bstrClickEventFunction：按钮点击后执行的回掉函数。
[id(18), helpstring("method AddToolButton")] HRESULT AddToolButton([in] BSTR bstrType, [in] BSTR bstrButtonID, [in] BSTR bstrButtonText, [in] BSTR bstrIconFileName, [in] BSTR bstrClickEventFunction);

//删除工具栏按钮：bstrButtonID：按钮ID；
[id(19), helpstring("method RemoveToolButton")] HRESULT RemoveToolButton([in] BSTR bstrType, [in] BSTR bstrButtonID);

//设置按钮状态。bstrState：可以是checked或者disabled
[id(20), helpstring("method SetToolButtonState")] HRESULT SetToolButtonState([in] BSTR bstrType, [in] BSTR bstrButtonID, [in] BSTR bstrState);

//获得按钮所在的屏幕位置。pbstrRect：返回坐标字符串，用英文逗号分割，分别是：left,top,right,bottom，例如100,100,200,120
[id(21), helpstring("method GetToolButtonRect")] HRESULT GetToolButtonRect([in] BSTR bstrType, [in] BSTR bstrButtonID, [out,retval] BSTR* pbstrRect);

//添加一个定时器：bstrCallback：回掉函数名称；nElapse：间隔
[id(22), helpstring("method AddTimer")] HRESULT AddTimer([in] BSTR bstrCallback, [in] LONG nElapse);

//删除一个定时器。
[id(23), helpstring("method RemoveTimer")] HRESULT RemoveTimer([in] BSTR bstrCallback);

//执行一个命令
[id(24), helpstring("method ExecCommand")] HRESULT ExecCommand([in] BSTR bstrCommandName, [in] BSTR bstrParams);
};
```
