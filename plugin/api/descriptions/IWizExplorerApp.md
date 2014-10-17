该对象是Wiz二次开发中最重要的一个对象，在插件html或者js脚本中，都可以直接访问。具体方法看这里：Wiz插件文档-整体框架

idl文件：WizExplorer.idl

```
[
object,
uuid(43C4CF61-5D88-4beb-B593-793A2251A491),
dual,
nonextensible,
helpstring("IWizExplorerApp Interface"),
pointer_default(unique)
]
interface IWizExplorerApp : IDispatch{
//获得当前打开的数据库。类型是：WizDatabase
[propget, id(1), helpstring("property Database")] HRESULT Database([out, retval] IDispatch** pVal);

//获得当前窗口对象。WizExplorerWindow
[propget, id(2), helpstring("property Window")] HRESULT Window([out, retval] IDispatch** pVal);

//获得Wiz安装路径，路径名后面包含反斜杠，例如：C:\Program files\WizBrother\Wiz\
[propget, id(3), helpstring("property AppPath")] HRESULT AppPath([out, retval] BSTR* pVal);

//获得当前使用的数据库存储文件夹，例如：D:\My Documents\My Knowledge\
[propget, id(4), helpstring("property DataStore")] HRESULT DataStore([out, retval] BSTR* pVal);

//获得当前使用的设置文件，通常给WizSettings对象使用，例如：D:\My Documents\My Knowledge\WizKM.xml
[propget, id(5), helpstring("property SettingsFileName")] HRESULT SettingsFileName([out, retval] BSTR* pVal);

//获得当前使用的log文件名，例如：D:\My Documents\My Knowledge\WizKM.log
[propget, id(6), helpstring("property LogFileName")] HRESULT LogFileName([out, retval] BSTR* pVal);

//获得一个本地化的插件的字符串。需要传入的参数：bstrPluginAppGUID：在plugin.ini文件里面定义的插件appguid； bstrStringName：字符串名称；返回：本地化的字符串
[id(7), helpstring("method LoadPluginString")] HRESULT LoadPluginString([in] BSTR bstrPluginAppGUID, [in] BSTR bstrStringName, [out, retval] BSTR* pVal);

//获得一个本地化的插件的字符串
[id(8), helpstring("method LoadPluginString2")] HRESULT LoadPluginString2([in] IDispatch* pHtmlDocumentDisp, [in] BSTR bstrStringName, [out, retval] BSTR* pVal);

//翻译字符串，使用Wiz内部的本地化功能，从英文翻译到本地语言
[id(9), helpstring("method TranslateString")] HRESULT TranslateString([in] BSTR bstrString, [out,retval] BSTR* pbstrRet);

//本地化当前html文件，仅用于插件类型为HtmlDialog类型的插件html文件，可以参考高级选项插件
[id(10), helpstring("method PluginLocalizeHtmlDialog")] HRESULT PluginLocalizeHtmlDialog([in] IDispatch* pHtmlDocumentDisp);

//获得当前插件的app guid，仅用于插件类型为HtmlDialog类型的插件html文件。可以被CurPluginAppGUID代替。
[id(11), helpstring("method GetPluginAppGUID")] HRESULT GetPluginAppGUID([in] IDispatch* pHtmlDocumentDisp, [out,retval] BSTR* pVal);

//获得当前插件的路径，仅用于插件类型为HtmlDialog类型的插件html文件，可以被CurPluginAppPath代替。
[id(12), helpstring("method GetPluginAppPath")] HRESULT GetPluginAppPath([in] IDispatch* pHtmlDocumentDisp, [out,retval] BSTR* pVal);

//创建Wiz对象。绿色版本无法直接创建COM对象，因此需要使用该方法来创建Wiz内部对象。
[id(13), helpstring("method CreateWizObject")] HRESULT CreateWizObject([in] BSTR bstrObjectName, [out,retval] IDispatch** ppVal);

//创建对象。因为HTML里面创建COM对象的时候，可能会有安全提示，因此可以使用这个方法来替换 new ActiveXObject("xxx");
[id(14), helpstring("method CreateActiveXObject")] HRESULT CreateActiveXObject([in] BSTR bstrObjectName, [out,retval] IDispatch** ppVal);

//获得当前插件的App GUID，适用于插件类型为HtmlDialog的html文件，或者类型为ExecuteScript的js文件
[propget, id(15), helpstring("property CurPluginAppGUID")] HRESULT CurPluginAppGUID([out, retval] BSTR* pVal);

//获得当前插件的GUID，适用于插件类型为HtmlDialog的html文件，或者类型为ExecuteScript的js文件
[propget, id(16), helpstring("property CurPluginGUID")] HRESULT CurPluginGUID([out, retval] BSTR* pVal);

//获得当前插件的路径，适用于插件类型为HtmlDialog的html文件，或者类型为ExecuteScript的js文件。对于全局插件js文件，可以使用GetPluginPathByScriptFileName方法来获得插件的路径。
[propget, id(17), helpstring("property CurPluginAppPath")] HRESULT CurPluginAppPath([out, retval] BSTR* pVal);

//翻译字符串，适用于插件类型为HtmlDialog的html文件，或者类型为ExecuteScript的js文件
[id(18), helpstring("method LoadString")] HRESULT LoadString([in] BSTR bstrStringName, [out, retval] BSTR* pVal);

//执行js代码，一般用于引用外部js文件并且执行。
[id(19), helpstring("method RunScriptCode")] HRESULT RunScriptCode([in] BSTR bstrScriptCode, [in] BSTR bstrScriptLanguage);

//从文件执行js脚本，一般用于引用外部js文件并且执行。
[id(20), helpstring("method RunScriptFile")] HRESULT RunScriptFile([in] BSTR bstrScriptFileName, [in] BSTR bstrScriptLanguage);

//从ini文件本地化一个字符串，需要指定ini文件
[id(21), helpstring("method LoadStringFromFile")] HRESULT LoadStringFromFile([in] BSTR bstrFileName, [in] BSTR bstrStringName, [out, retval] BSTR* pVal);

//本地化一个html文件，需要指定语言ini文件
[id(22), helpstring("method LocalizeHtmlDocument")] HRESULT LocalizeHtmlDocument([in] BSTR bstrLanguageFileName, [in] IDispatch* pHtmlDocumentDisp);

//获得html路径。通过传入一个html的document对象，获得该html对应得文件名（包含路径）。
//例如：html的url是： file://d:/my%20documents/my%20knowledge/plugins/xxx/a.htm，
//var pluginpath = objApp.GetHtmlDocumentPath(document);
//返回：D:\My Documents\My Knowledge\Plugins\xxx\a.htm
//要求html文件必须是本地的
[id(23), helpstring("method GetHtmlDocumentPath")] HRESULT GetHtmlDocumentPath([in] IDispatch* pHtmlDocumentDisp, [out,retval] BSTR* pVal);

//通过一个js文件名，获得插件所在的路径。该js文件名不需要包含路径。
//例如：var pluginpath = objApp.GetPluginPathByScriptFileName("km.js");
[id(24), helpstring("method GetPluginPathByScriptFileName")] HRESULT GetPluginPathByScriptFileName([in] BSTR bstrScriptFileName, [out,retval] BSTR* pbstrPluginPath);

//添加全局脚本。该脚本在Wiz运行期间，一直有效。
[id(25), helpstring("method AddGlobalScript")] HRESULT AddGlobalScript([in] BSTR bstrScriptFileName);

//执行一个插件。传入一个guid，可以执行该guid对应得插件。注意：仅能执行脚本类型是HtmlDialog或者ExecuteScript类型的插件。
[id(26), helpstring("method ExecutePlugin")] HRESULT ExecutePlugin([in] BSTR bstrPluginGUID);

//执行一个命令
[id(27), helpstring("method ExecuteCommand")] HRESULT ExecuteCommand([in] BSTR bstrCommandName, [in] VARIANT* pvInParam1, [in] VARIANT* pvInParam2, [out, retval] VARIANT* pvretResult);

//插入插件菜单
[id(28), helpstring("method InsertPluginMenu")] HRESULT InsertPluginMenu([in] LONGLONG nMenuHandle, [in] LONG nInsertPos, [in] BSTR bstrMenuType, [in] LONG nPluginBeginCommand, [out, retval] BSTR* pbstrCommandPluginGUID);

//获得同步进度对话框，返回IWizSyncProgressDlg
[propget, id(29), helpstring("property SyncProgress")] HRESULT SyncProgress([out, retval] IDispatch** pVal);

//获得群组数据库对象，传入群组KBGUID，返回IWizDatabase
[id(30), helpstring("method GetGroupDatabase")] HRESULT GetGroupDatabase([in] BSTR bstrKbGUID, [out, retval] IDispatch** pVal);
};
```

WizExplorerApp对象，主要用途就是对外提供了Wiz的一些内部接口，同时提供插件的本地化等功能。

对于不同的插件类型，获得WizExplorerApp对象的方式也不同。

对于ExecuteScript类型的插件，该类插件功能对应的是一个js或者vbs文件。Wiz在执行这类脚本的时候，内置了一个对象，名称就是WizExplorerApp。因此您可以直接使用。

对于HtmlDialog类型的插件，html文件里面的window.external对象，就是WizExplorerApp。

对于全局插件，js脚本里面，可以直接使用objApp对象（已经是WizExplorerApp了）。

这三种类型的插件，可以参考[例子](../../file/wizsamples.zip)
