IWizSettings是WizKMCore.dll包含的一个COM对象。IWizSettings对象独立于数据库，因此您可以直接创建这个对象。

IWizSettings主要用来操作Wiz用户存储目录的WizKM.xml文件中，该文件名，可以通过WizExploer.WizExploerApp对象来获得这个文件名

和IWizMeta不同，IWizSettings是保存在独立的Xml文件中，和数据库无关，是全局的设置。

在Wiz自带的朗读插件里面，有这个对象的使用方法：

```
    var objApp = new ActiveXObject("WizExplorer.WizExplorerApp");
    var objSettings = new ActiveXObject("WizKMCore.WizSettings");
    objSettings.Open(objApp.SettingsFileName);        //打开用户设置文件
    ...
    var initName = objSettings.StringValue("Plugin_Speech", "VocieName");    //从设置里面获得用户信息
    ...
    objSettings.StringValue("Plugin_Speech", "VocieName") = selectVoiceName.value;        //保存用户信息到设置文件里面
```

ProgID：WizKMCore.WizSettings

文件：WizKMCore.dll

```
[
    object,
    uuid(E3C896AD-3E1F-453E-800C-07A1837E2F23),
    dual,
    nonextensible,
    helpstring("IWizSettings Interface"),
    pointer_default(unique)
]
interface IWizSettings : IDispatch{

    //打开设置文件
    [id(1), helpstring("method Open")] HRESULT Open([in] BSTR bstrFileName);

    //关闭设置文件
    [id(2), helpstring("method Close")] HRESULT Close(void);

    //判断设置文件是否已经修改了
    [propget, id(3), helpstring("property IsDirty")] HRESULT IsDirty([out, retval] VARIANT_BOOL* pVal);

    //获得/设置一个字符串信息
    [propget, id(4), helpstring("property StringValue")] HRESULT StringValue([in] BSTR bstrSection, [in] BSTR bstrKey, [out, retval] BSTR* pVal);
    [propput, id(4), helpstring("property StringValue")] HRESULT StringValue([in] BSTR bstrSection, [in] BSTR bstrKey, [in] BSTR newVal);

    //获得/设置一个整数信息
    [propget, id(5), helpstring("property IntValue")] HRESULT IntValue([in] BSTR bstrSection, [in] BSTR bstrKey, [out, retval] LONG* pVal);
    [propput, id(5), helpstring("property IntValue")] HRESULT IntValue([in] BSTR bstrSection, [in] BSTR bstrKey, [in] LONG newVal);

    //获得/设置一个布尔值
    [propget, id(6), helpstring("property BoolValue")] HRESULT BoolValue([in] BSTR bstrSection, [in] BSTR bstrKey, [out, retval] VARIANT_BOOL* pVal);
    [propput, id(6), helpstring("property BoolValue")] HRESULT BoolValue([in] BSTR bstrSection, [in] BSTR bstrKey, [in] VARIANT_BOOL newVal);

    //获得一个Section，返回值是IWizSettingsSection对象
    [propget, id(7), helpstring("property Section")] HRESULT Section([in] BSTR bstrSection, [out, retval] IDispatch** pVal);

    //删除一个Section
    [id(8), helpstring("method ClearSection")] HRESULT ClearSection([in] BSTR bstrSection);
};
```
