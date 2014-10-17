在编写插件的时候，经常要用到配置文件来保存/恢复一些状态，因此配置文件读写就是必须的了。

在为知笔记里面，配置文件通常有两类，一种是和用户帐号相关的数据，例如不同的帐号，配置不同。还有一种是和帐号无关，所有的帐号使用相同的配置。

### 和帐号相关的配置

可以使用IWizDatabase里面的Meta来保存。例如下面的代码：

```
//从数据库读取数据
var settings_meta = "MyPlugin";
textAddress.value = objDatabase.GetMeta(settings_meta, "Address");
textPort.value = objDatabase.GetMeta(settings_meta, "Port");
textUserName.value = objDatabase.GetMeta(settings_meta, "UserName");
textPassword.value = objDatabase.GetMeta(settings_meta, "Password");
textImageWidth.value = objDatabase.GetMeta(settings_meta, "ImageWidth");
checkAutoAddMoreTag.checked = objDatabase.GetMeta(settings_meta, "AutoAddMoreTag") == "1";
textCategory.value = objDatabase.GetMeta(settings_meta, "Category");
selectSystem.value = objDatabase.GetMeta(settings_meta, "System");
checkUseSEO.checked = objDatabase.GetMeta(settings_meta, "UseSEO") == "1";
checkUseTime.checked = objDatabase.GetMeta(settings_meta, "UseTime") == "1";
checkTagAsCategory.checked = objDatabase.GetMeta(settings_meta, "TagAsCategory") == "1";

//保存设置到数据库
//save settings to objDatabase
objDatabase.SetMeta(settings_meta, "Port", port);
objDatabase.SetMeta(settings_meta, "UserName", username);
objDatabase.SetMeta(settings_meta, "Password", password);
objDatabase.SetMeta(settings_meta, "ImageWidth", textImageWidth.value);
objDatabase.SetMeta(settings_meta, "AutoAddMoreTag", checkAutoAddMoreTag.checked ? "1" : "0");
objDatabase.SetMeta(settings_meta, "Category", textCategory.value);
objDatabase.SetMeta(settings_meta, "System", blog_system);
objDatabase.SetMeta(settings_meta, "UseSEO", checkUseSEO.checked ? "1" : "0");
objDatabase.SetMeta(settings_meta, "UseTime", checkUseTime.checked ? "1" : "0");
objDatabase.SetMeta(settings_meta, "TagAsCategory", checkTagAsCategory.checked ? "1" : "0");
```

这些数据，将会被保存在index.db里面，每一个帐号都是独立的数据。

### 和帐号无关的配置

可以把数据保存在Wiz.xml文件里面（用户数据存储路径下面）。

```
//创建配置文件读写对象
var objSettings = objApp.CreateWizObject("WizKMCore.WizSettings");
//打开默认的配置文件
objSettings.Open(objApp.SettingsFileName);

//从设置文件里面读取数据
textFolder.value = objSettings.GetStringValue(settings_section, "Folder");
buttonUseTemplate.checked = objSettings.GetBoolValue(settings_section, "UseTemplate");
buttonSEOAsFileName.checked = objSettings.GetBoolValue(settings_section, "SEOAsFileName");
selectCharset.selectedIndex = objSettings.GetIntValue(settings_section, "Charset");

//保存设置
objSettings.SetStringValue(settings_section, "Folder", textFolder.value);
objSettings.SetBoolValue(settings_section, "UseTemplate", buttonUseTemplate.checked);
objSettings.SetBoolValue(settings_section, "SEOAsFileName", buttonSEOAsFileName.checked);
objSettings.SetIntValue(settings_section, "Charset", selectCharset.selectedIndex);
objSettings.Close();
```

### 注册表读写

使用IWizCommonUI对象

```
var objCommon = objApp.CreateWizObject("WizKMControls.WizCommonUI");
```


```
//从注册表中读入一个值。bstrRoot：注册表根键，可以是 HKEY_LOCAL_MACHINE，HKEY_CLASSES_ROOT，HKEY_USERS，HKEY_CURRENT_CONFIG；bstrKey：注册表键值路径；bstrValueName：注册表的值名称；返回值：注册表的值
[id(29), helpstring("method GetValueFromReg")] HRESULT GetValueFromReg([in] BSTR bstrRoot, [in] BSTR bstrKey, [in] BSTR bstrValueName, [out,retval] BSTR* pbstrData);

//向注册表写入一个值。bstrRoot：注册表根键，可以是 HKEY_LOCAL_MACHINE，HKEY_CLASSES_ROOT，HKEY_USERS，HKEY_CURRENT_CONFIG；bstrKey：注册表键值路径；bstrValueName：注册表的值名称；bstrData：数据值；bstrDataType：数据类型，可以是REG_SZ或者REG_DWORD。
[id(30), helpstring("method SetValueToReg")] HRESULT SetValueToReg([in] BSTR bstrRoot, [in] BSTR bstrKey, [in] BSTR bstrValueName, [in] BSTR bstrData, [in] BSTR bstrDataType);

//枚举注册表的值名称。bstrRoot：注册表根键，可以是 HKEY_LOCAL_MACHINE，HKEY_CLASSES_ROOT，HKEY_USERS，HKEY_CURRENT_CONFIG；bstrKey：注册表键值路径；返回值：所有的值，类型为安全数组，如果在javascript里面使用，请参阅本文后面部分。
[id(31), helpstring("method EnumRegValues")] HRESULT EnumRegValues([in] BSTR bstrRoot, [in] BSTR bstrKey, [out, retval] VARIANT* pvValueNames);

//枚举注册表的键名称。bstrRoot：注册表根键，可以是 HKEY_LOCAL_MACHINE，HKEY_CLASSES_ROOT，HKEY_USERS，HKEY_CURRENT_CONFIG；bstrKey：注册表键值路径；返回值：所有的key名称，类型为安全数组，如果在javascript里面使用，请参阅本文后面部分。
[id(32), helpstring("method EnumRegKeys")] HRESULT EnumRegKeys([in] BSTR bstrRoot, [in] BSTR bstrKey, [out, retval] VARIANT* pvKeys);
```

### ini文件读写

```
var objCommon = objApp.CreateWizObject("WizKMControls.WizCommonUI");
```

```
//从ini文件读取一个字符串。bstrFileName：ini文件名；bstrSection：ini里面的section；bstrKey：key；返回值：读取到的值
[id(21), helpstring("method GetValueFromIni")] HRESULT GetValueFromIni([in] BSTR bstrFileName, [in] BSTR bstrSection, [in] BSTR bstrKey, [out,retval] BSTR* pbstrValue);

//向ini文件写入一个字符串。bstrFileName：ini文件名；bstrSection：ini里面的section；bstrKey：key；bstrValue：写入的值
[id(22), helpstring("method SetValueToIni")] HRESULT SetValueToIni([in] BSTR bstrFileName, [in] BSTR bstrSection, [in] BSTR bstrKey, [in] BSTR bstrValue);

//获得ini文件所有的section
[id(64), helpstring("method GetIniSections")] HRESULT GetIniSections([in] BSTR bstrFileName, [out, retval] VARIANT* pvValueNames);
```

通过三、四两种方式，可以自己处理保存的数据，例如保存位置，文件等等。
