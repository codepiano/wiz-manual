IWizIECapture是NPWizWebCapture.dll包含的一个COM对象。IWizIECapture主要用来保存网页，例如在IE里面保存网页。

IWizIECapture可以在浏览器内程序调用，让您的浏览器程序很容易的增加网页保存功能。

ProgID：WizWebCapture.WizIECapture

文件：NPWizWebCapture.dll

```
[
    object,
    uuid(502644EE-7301-457B-A77D-8BC3329AEA21),
    dual,
    nonextensible,
    helpstring("IWizIECapture Interface"),
    pointer_default(unique)
]
interface IWizIECapture : IDispatch{

    //保存一个网页。pHtmlDocument2Disp：IHTMLDocument2对象；vbSaveSel：是否保存选中部分；返回值：保存成功的IWizDocument对象
    [id(1), helpstring("method Save")] HRESULT Save([in] IDispatch* pHtmlDocument2Disp, [in] VARIANT_BOOL vbSaveSel, [out,retval] IDispatch** ppWizDocumentDisp);

    //在后台保存网页，避免一直占用浏览器前台，让用户无法操作。pHtmlDocument2Disp：IHTMLDocument2对象；vbSaveSel：是否保存选中部分。
    [id(2), helpstring("method BackgroundSave")] HRESULT BackgroundSave([in] IDispatch* pHtmlDocument2Disp, [in] VARIANT_BOOL vbSaveSel);

    //获得或者设置数据库路径
    [propget, id(3), helpstring("property DatabasePath")] HRESULT DatabasePath([out, retval] BSTR* pVal);
    [propput, id(3), helpstring("property DatabasePath")] HRESULT DatabasePath([in] BSTR newVal);
};


在Wiz安装目录下面的save.htm文件中，有这个对象的使用方法：

function OnContextMenu()
{
    var objCapture= null;
    //
    try
    {
        objCapture = new ActiveXObject("WizWebCapture.WizIECapture");
    }
    catch ( e)
    {
        alert("WizBrother WebCapture haven't installed in your computer!");
    }
    //
    try
    {
        objCapture.BackgroundSave(external.menuArguments.document, false);
    }
    catch (e)
    {
        alert("Failed to save the web page!");
    }
}


OnContextMenu();
```
