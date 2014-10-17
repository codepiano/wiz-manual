IWizCommonUI是WizKMControls.dll包含的一个COM对象。IWizCommonUI主要提供了一些常用的对话框，界面操作，二次开发中常用的功能等等。

IWizCommonUI提供了一些Wiz的一些常用的对话框，例如获得用户信息，添加文档到日历对话框，选项对话框，新建/编辑日历事件对话框，关于对话框等等。

IWizCommonUI还提供了一些通用的对话框，例如提示用户输入一个整数，输入一段文字，选择一个文件夹或者文件等等。

IWizCommonUI还提供了一些常用的操作，例如获得一个文件夹下面的文件或者子文件夹，注册表操作等等。

ProgID：WizKMControls.WizCommonUI

文件：WizKMControls.dll

```
[
    object,
    uuid(B671315F-B6FE-4799-AE50-AEA0848FBA05),
    dual,
    nonextensible,
    helpstring("IWizCommonUI Interface"),
    pointer_default(unique)
]
interface IWizCommonUI : IDispatch{

    //显示一个对话框，请求用户输入用户名和密码。hWnd：当前窗口；pbstrUserName：返回用户输入的用户名；pbstrPassword：返回用户输入的密码
    [id(1), helpstring("method QueryUserAccount")] HRESULT QueryUserAccount([in] LONGLONG hWnd, [in, out] BSTR* pbstrUserName, [in, out] BSTR* pbstrPassword);

    //添加文档到日历中。这个方法会出现一个对话框，提示用户输入参数。pDocumentDisp：文档对象，类型为IWizDocument；pbRet：用户是否点击了确定按钮
    [id(2), helpstring("method AddDocumentToCalendar")] HRESULT AddDocumentToCalendar([in] IDispatch* pDocumentDisp, [out,retval] VARIANT_BOOL* pbRet);

    //创建一个日历事件。显示一个对话框，提示用户输入某些参数。pDatabaseDisp：数据库对象；dtEvent：事件开始时间；ppDocumentDisp：成功创建事件后，自动生成的文档
    [id(3), helpstring("method CreateCalendarEvent")] HRESULT CreateCalendarEvent([in] IDispatch* pDatabaseDisp, [in] DATE dtEvent, [out,retval] IDispatch** ppDocumentDisp);

    //编辑日历事件。显示一个对话框，提示用户编辑。pDocumentDisp：文档对象，类型为IWizDocument；pbRet：用户是否点击了确定按钮
    [id(4), helpstring("method EditCalendarEvent")] HRESULT EditCalendarEvent([in] IDispatch* pDocumentDisp, [out,retval] VARIANT_BOOL* pbRet);

    //显示关于对话框。bstrModuleName：EXE文件名
    [id(5), helpstring("method AboutBox")] HRESULT AboutBox([in] BSTR bstrModuleName);

    //显示选项对话框。nFlags：保留，必须为0
    [id(6), helpstring("method OptionsDlg")] HRESULT OptionsDlg([in] LONG nFlags);

    //显示备份对话框。nFlags：保留，必须为0
    [id(7), helpstring("method BackupDlg")] HRESULT BackupDlg([in] LONG nFlags);

    //显示帐户管理对话框。nFlags：保留，必须为0
    [id(8), helpstring("method AccountsManagerDlg")] HRESULT AccountsManagerDlg([in] LONG nFlags);

    //显示选择帐户对话框。nFlags：保留，必须为0；pbstrDatabasePath：用户选择的帐户对应的数据库路径。
    [id(9), helpstring("method ChooseAccount")] HRESULT ChooseAccount([in] LONG nFlags, [out,retval] BSTR* pbstrDatabasePath);

    //显示一个对话框，让用户输入一个整数。bstrTitle：对话框标题；bstrDescription：描述文字；nInitValue：默认值，用来初始化对话框；pvbOK：用户是否点击了确定按钮；pnRet：返回用户输入的结果
    [id(10), helpstring("method GetIntValue")] HRESULT GetIntValue([in] BSTR bstrTitle, [in] BSTR bstrDescription, [in] LONG nInitValue, [out] VARIANT_BOOL* pvbOK, [out,retval] LONG* pnRet);

    //显示一个对话框，让用户输入一个整数。bstrTitle：对话框标题；bstrDescription：描述文字；nInitValue：默认值，用来初始化对话框；nMin：最小值；nMax：最大值；pvbOK：用户是否点击了确定按钮；pnRet：返回用户输入的结果
    [id(11), helpstring("method GetIntValue2")] HRESULT GetIntValue2([in] BSTR bstrTitle, [in] BSTR bstrDescription, [in] LONG nInitValue, [in] LONG nMin, [in] LONG nMax, [out] VARIANT_BOOL* pvbOK, [out,retval] LONG* pnRet);

    //检查新Wiz版本。vbPrompt：是否提示用户是否存在新版本；vbSilent：是否安静安装
    [id(12), helpstring("method CheckNewVersion")] HRESULT CheckNewVersion([in] VARIANT_BOOL vbPrompt, [in] VARIANT_BOOL vbSilent);

    //显示一个对话框，提示用户输入一行文字。bstrTitle：标题；bstrDescription：描述；bstrInitValue：初始值；返回值：返回用户录入的结果
    [id(13), helpstring("method InputBox")] HRESULT InputBox([in] BSTR bstrTitle, [in] BSTR bstrDescription, [in] BSTR bstrInitValue, [out,retval] BSTR* pbstrText);

    //显示一个对话框，提示用户输入多行文字。bstrTitle：标题；bstrDescription：描述；bstrInitValue：初始值；返回值：返回用户录入的结果
    [id(14), helpstring("method InputMultiLineText")] HRESULT InputMultiLineText([in] BSTR bstrTitle, [in] BSTR bstrDescription, [in] BSTR bstrInitValue, [out,retval] BSTR* pbstrText);

    //提示用户选择一个磁盘文件。vbOpen：打开文件或者保存文件；bstrFilter：文件名过滤器，格式为：Html文件(*.htm;*.html)|*.htm;*.html|所有文件(*.*)|*.*|；返回值：用户选择的文件
    [id(15), helpstring("method SelectWindowsFile")] HRESULT SelectWindowsFile([in] VARIANT_BOOL vbOpen, [in] BSTR bstrFilter, [out,retval] BSTR* pbstrResultFileName);

    //提示用户选择一个磁盘文件夹。bstrDescription：描述；返回值：用户选择的文件夹
    [id(16), helpstring("method SelectWindowsFolder")] HRESULT SelectWindowsFolder([in] BSTR bstrDescription, [out,retval] BSTR* pbtrFolderPath);

    //选择一个Wiz文件夹。bstrDatabasePath：数据库路径；bstrDescription：描述；返回值：用户选择的文件夹对象，类型为IWizFolder 
    [id(17), helpstring("method SelectWizFolder")] HRESULT SelectWizFolder([in] BSTR bstrDatabasePath, [in] BSTR bstrDescription, [out,retval] IDispatch** ppFolderDisp);

    //选择一个Wiz文档。bstrDatabasePath：数据库路径；bstrDescription：描述；返回值：用户选择的文档对象，类型为IWizDocument
    [id(18), helpstring("method SelectWizDocument")] HRESULT SelectWizDocument([in] BSTR bstrDatabasePath, [in] BSTR bstrDescription, [out,retval] IDispatch** ppDocumentDisp);

    //从一个text文件获得文字内容。可以有效地获得避免乱码。bstrFileName：文件名；返回值：text里面的文字
    [id(19), helpstring("method LoadTextFromFile")] HRESULT LoadTextFromFile([in] BSTR bstrFileName, [out,retval] BSTR* pbstrText);

    //保存文字为text文件。bstrFileName：文件名；bstrText：文字；bstrCharset：字符集，可以是unicode（包含签名），utf-8（没有bom签名），utf-8-bom（utf-8编码，保存的文件里面包含bom签名），或者其他的字符集，例如gb2312, gbk, big5等等
    [id(20), helpstring("method SaveTextToFile")] HRESULT SaveTextToFile([in] BSTR bstrFileName, [in] BSTR bstrText, [in] BSTR bstrCharset);

    //从ini文件读取一个字符串。bstrFileName：ini文件名；bstrSection：ini里面的section；bstrKey：key；返回值：读取到的值
    [id(21), helpstring("method GetValueFromIni")] HRESULT GetValueFromIni([in] BSTR bstrFileName, [in] BSTR bstrSection, [in] BSTR bstrKey, [out,retval] BSTR* pbstrValue);

    //向ini文件写入一个字符串。bstrFileName：ini文件名；bstrSection：ini里面的section；bstrKey：key；bstrValue：写入的值
    [id(22), helpstring("method SetValueToIni")] HRESULT SetValueToIni([in] BSTR bstrFileName, [in] BSTR bstrSection, [in] BSTR bstrKey, [in] BSTR bstrValue);

    //生成一个临时文件名。bstrFileExt：扩展名，例如.txt；返回值：生成的临时文件名，包含全路径。
    [id(23), helpstring("method GetATempFileName")] HRESULT GetATempFileName([in] BSTR bstrFileExt, [out,retval] BSTR* pbstrFileName);

    //获得指定的文件夹路径。bstrFolderName：文件夹名称，可能的值为：
    //WindowsFolder，Windows安装目录，一般是C:\Windows
    //SystemFolder：系统目录，C:\Windows\System32
    //TemporaryFolder：临时文件夹
    //AppPath：Wiz安装路径
    //返回值：全路径名
    [id(24), helpstring("method GetSpecialFolder")] HRESULT GetSpecialFolder([in] BSTR bstrFolderName, [out,retval] BSTR* pbstrPath);

    //执行一个exe文件。bstrExeFileName：EXE文件名；bstrParams：参数；vbWait：是否等待EXE执行完毕再返回；返回值：EXE返回结果，只有vbWait=true的时候才有效
    [id(25), helpstring("method RunExe")] HRESULT RunExe([in] BSTR bstrExeFileName, [in] BSTR bstrParams, [in] VARIANT_BOOL vbWait, [out,retval] long* pnRet);

    //枚举一个文件夹下面的文件。bstrPath：文件夹路径；bstrFileExt：扩展名，格式为*.txt;*.jpg;*.png或者*.*；vbIncludeSubFolders：是否包含子文件夹；返回值：枚举到的文件名。类型为一个安全数组，如果在javascript里面使用，请参阅本文后面部分。
    [id(26), helpstring("method EnumFiles")] HRESULT EnumFiles([in] BSTR bstrPath, [in] BSTR bstrFileExt, [in] VARIANT_BOOL vbIncludeSubFolders, [out,retval] VARIANT* pvFiles);

    ////枚举一个文件夹下面的文件夹。bstrPath：文件夹路径；返回值：枚举到的文件名。类型为一个安全数组，如果在javascript里面使用，请参阅本文后面部分
    [id(27), helpstring("method EnumFolders")] HRESULT EnumFolders([in] BSTR bstrPath, [out,retval] VARIANT* pvFolders);

    //创建一个COM对象。因为在html里面用javascript的 new ActiveXObject方式创建COM对象，可能会有安全警告，因此可以使用这个方法来代替，避免出现安全警告。
    //bstrProgID：COM对象的ProgID。返回值：成功创建的对象。
    [id(28), helpstring("method CreateActiveXObject")] HRESULT CreateActiveXObject([in] BSTR bstrProgID, [out,retval] IDispatch** ppObjDisp);

    //从注册表中读入一个值。bstrRoot：注册表根键，可以是 HKEY_LOCAL_MACHINE，HKEY_CLASSES_ROOT，HKEY_USERS，HKEY_CURRENT_CONFIG；bstrKey：注册表键值路径；bstrValueName：注册表的值名称；返回值：注册表的值
    [id(29), helpstring("method GetValueFromReg")] HRESULT GetValueFromReg([in] BSTR bstrRoot, [in] BSTR bstrKey, [in] BSTR bstrValueName, [out,retval] BSTR* pbstrData);

    //向注册表写入一个值。bstrRoot：注册表根键，可以是 HKEY_LOCAL_MACHINE，HKEY_CLASSES_ROOT，HKEY_USERS，HKEY_CURRENT_CONFIG；bstrKey：注册表键值路径；bstrValueName：注册表的值名称；bstrData：数据值；bstrDataType：数据类型，可以是REG_SZ或者REG_DWORD。
    [id(30), helpstring("method SetValueToReg")] HRESULT SetValueToReg([in] BSTR bstrRoot, [in] BSTR bstrKey, [in] BSTR bstrValueName, [in] BSTR bstrData, [in] BSTR bstrDataType);

    //枚举注册表的值名称。bstrRoot：注册表根键，可以是 HKEY_LOCAL_MACHINE，HKEY_CLASSES_ROOT，HKEY_USERS，HKEY_CURRENT_CONFIG；bstrKey：注册表键值路径；返回值：所有的值，类型为安全数组，如果在javascript里面使用，请参阅本文后面部分。
    [id(31), helpstring("method EnumRegValues")] HRESULT EnumRegValues([in] BSTR bstrRoot, [in] BSTR bstrKey, [out, retval] VARIANT* pvValueNames);

    //枚举注册表的键名称。bstrRoot：注册表根键，可以是 HKEY_LOCAL_MACHINE，HKEY_CLASSES_ROOT，HKEY_USERS，HKEY_CURRENT_CONFIG；bstrKey：注册表键值路径；返回值：所有的key名称，类型为安全数组，如果在javascript里面使用，请参阅本文后面部分。
    [id(32), helpstring("method EnumRegKeys")] HRESULT EnumRegKeys([in] BSTR bstrRoot, [in] BSTR bstrKey, [out, retval] VARIANT* pvKeys);

    //选择一个Wiz模板。bstrTemplateType：模板类型，可以是new或者display；返回值：用户选择的模板文件名。
    [id(33), helpstring("method SelectTemplate")] HRESULT SelectTemplate([in] BSTR bstrTemplateType, [out, retval] BSTR* pbstrResultTemplateFileName);

    //从HTML标记里面获得一个属性值。bstrHtmlTag：HTML标记内容；bstrTagAttributeName：属性名；返回值：属性值
    [id(34), helpstring("method HtmlTagGetAttributeValue")] HRESULT HtmlTagGetAttributeValue([in] BSTR bstrHtmlTag, [in] BSTR bstrTagAttributeName, [out, retval] BSTR* pbstrAttributeValue);

    //从HTML文字里面提取一个或者多个标记。bstrHtmlText：HTML文字；bstrTagName：HTML标记名；bstrTagAttributeName：HTML标记属性名；bstrTagAttributeValue：HTML标记属性值；返回值：所有符合条件的标记，类型为安全数组。如果在javascript里面使用，请参阅本文后面部分。
    [id(35), helpstring("method HtmlExtractTags")] HRESULT HtmlExtractTags([in] BSTR bstrHtmlText, [in] BSTR bstrTagName, [in] BSTR bstrTagAttributeName, [in] BSTR bstrTagAttributeValue, [out, retval] VARIANT* pvTags);

    //从HTML里面获得所有的链接。bstrHtmlText：HTML文字；bstrURL：HTML的URL；返回值：所有链接，类型为安全数组。如果在javascript里面使用，请参阅本文后面部分。
    [id(36), helpstring("method HtmlEnumLinks")] HRESULT HtmlEnumLinks([in] BSTR bstrHtmlText, [in] BSTR bstrURL, [out, retval] VARIANT* pvLinks);

    //从HTML文字里面，提取HTML正文。bstrHtmlText：HTML文字；bstrURL：HTML的URL；返回值：HTML正文
    [id(37), helpstring("method HtmlGetContent")] HRESULT HtmlGetContent([in] BSTR bstrHtmlText, [in] BSTR bstrURL, [out, retval] BSTR* pbstrContent);

    //转换一个文件为HTML文件。bstrSrcFileName：原文件；返回值：转换后生成的HTML文件
    [id(38), helpstring("method HtmlConvertFileToHtmlFile")] HRESULT HtmlConvertFileToHtmlFile([in] BSTR bstrSrcFileName, [out, retval] BSTR* pbstrHtmlFileName);

    //将一个zip文件(ziw文件)保存成一个html文件。bstrSrcZipFileName：zip文件名；bstrDestHtmlFileName：目标HTML文件；bstrHtmlTitle：HTML文件标题。
    [id(39), helpstring("method HtmlConvertZipFileToHtmlFile")] HRESULT HtmlConvertZipFileToHtmlFile([in] BSTR bstrSrcZipFileName, [in] BSTR bstrDestHtmlFileName, [in] BSTR bstrHtmlTitle);

    //显示选项对话框

    [id(40), helpstring("method OptionsDlg2")] HRESULT OptionsDlg2([in] IDispatch* pDisp, [in] LONG nFlags);

    //显示一个对话框，选择一个日期。
    [id(41), helpstring("method SelectDay")] HRESULT SelectDay([in] DATE dtDefault, [out, retval] DATE* pRet);

    //报告同步结果
    [id(42), helpstring("method ReportSyncResult")] HRESULT ReportSyncResult([in] LONGLONG hWnd, [in] VARIANT vDownloaded, [in] VARIANT vUploaded, [in] VARIANT vSkipped, [in] VARIANT vFailed);

    //复制指定的文字到剪贴板
    [id(43), helpstring("method CopyTextToClipboard")] HRESULT CopyTextToClipboard([in] BSTR bstrText);

    //提示流量达到限制
    [id(44), helpstring("method PromptTrafficLimit")] HRESULT PromptTrafficLimit([in] IDispatch* pDatabaseDisp, [in] LONGLONG hWnd, [in] LONGLONG nLimit);

    //显示帐户安全对话框
    [id(45), helpstring("method AccountSecurityDlg")] HRESULT AccountSecurityDlg([in] IDispatch* pDisp, [in] LONG nFlags);

    //显示初始化加密向导对话框
    [id(46), helpstring("method InitEncryptionDlg")] HRESULT InitEncryptionDlg([in] IDispatch* pDisp, [in] LONG nFlags);

    //显示导入证书对话框
    [id(47), helpstring("method ImportCertDlg")] HRESULT ImportCertDlg([in] IDispatch* pDisp, [in] LONG nFlags);

    //提示用户输入一个密码
    [id(48), helpstring("method EnterPasswordDlg")] HRESULT EnterPasswordDlg([in] BSTR bstrTitle, [in] BSTR bstrInfo, [in] VARIANT_BOOL vbShowRememberPasswordButton, [out] VARIANT_BOOL* pvbRememberPassword, [out, retval] BSTR* pbstrPassword);

    //请求用户账号
    [id(49), helpstring("method QueryUserAccount2")] HRESULT QueryUserAccount2([in] IDispatch* pDatabaseDisp, [in] LONGLONG hWnd, [in, out] BSTR* pbstrUserName, [in, out] BSTR* pbstrPassword);

    //将剪贴板内保存的图片保存成图片文件
    [id(50), helpstring("method ClipboardToImage")] HRESULT ClipboardToImage([in] OLE_HANDLE hwnd, [in] BSTR bstrOptions, [out,retval] BSTR* pbstrImageFileName);

    //抓取屏幕并获得最后的图片文件
    [id(51), helpstring("method CaptureScreen")] HRESULT CaptureScreen([in] BSTR bstrOptions, [out,retval] BSTR* pbstrImageFileName);

    //分解chm文件，还没有实现
    [id(52), helpstring("method DumpCHM")] HRESULT DumpCHM([in] BSTR bstrCHMFileName, [in] BSTR bstrTempPath, [in] BSTR bstrOptions, [out,retval] VARIANT* pvRet);

    //显示插入超级链接对话框
    [id(53), helpstring("method HyperlinkDlg")] HRESULT HyperlinkDlg([in] IDispatch* pdispDatabase, [in] IDispatch* pdispDocument, [in] IDispatch* pdispHtmlDocument, [in] BSTR bstrInitURL, [out,retval] BSTR* pbstrResultURL);

    //选择一个颜色
    [id(54), helpstring("method SelectColor")] HRESULT SelectColor([in] BSTR bstrInitColor, [out,retval] BSTR* pbstrRetColor);

    //获得系统里面的字体
    [id(55), helpstring("method GetFontNames")] HRESULT GetFontNames([out,retval] VARIANT* pvNames);

    //选择一个字体。
    [id(56), helpstring("method SelectFont")] HRESULT SelectFont([in] BSTR bstrInit, [out,retval] BSTR* pbstrRet);

    //创建一个Wiz内部对象        

    [id(57), helpstring("method CreateWizObject")] HRESULT CreateWizObject([in] BSTR bstrObjectName, [out,retval] IDispatch** ppVal);

    //获得脚本参数，用于在执行脚本的时候使用
    [id(58), helpstring("method GetScriptParamValue")] HRESULT GetScriptParamValue([in] BSTR bstrParamName, [out,retval] BSTR* pbstrVal);

    //从ini文件里面获得一个本地化的字符串
    [id(59), helpstring("method LoadStringFromFile")] HRESULT LoadStringFromFile([in] BSTR bstrFileName, [in] BSTR bstrStringName, [out, retval] BSTR* pVal);

    //给定一个URL，下载该URL并获得文件的文本（必须是文本文件）
    [id(60), helpstring("method URLDownloadToText")] HRESULT URLDownloadToText([in] BSTR bstrURL, [out, retval] BSTR* pbstrText);

    //选择一个或者多个Wiz标签
    [id(61), helpstring("method SelectWizTags")] HRESULT SelectWizTags([in] BSTR bstrDatabasePath, [in] BSTR bstrTitle, [in] LONG nFlags, [out,retval] IDispatch** ppTagCollectionDisp);

    //显示日历对话框
    [id(62), helpstring("method CalendarDlg")] HRESULT CalendarDlg([in] LONG nFlags);

    //显示帐户信息对话框
    [id(63), helpstring("method AccountInfoDlg")] HRESULT AccountInfoDlg([in] IDispatch* pDatabaseDisp, [in] LONG nFlags);

    //获得ini文件所有的section
    [id(64), helpstring("method GetIniSections")] HRESULT GetIniSections([in] BSTR bstrFileName, [out, retval] VARIANT* pvValueNames);

    //获得ini里面的一个本地化字符串
    [id(65), helpstring("method LoadStringFromFile2")] HRESULT LoadStringFromFile2([in] BSTR bstrFileName, [in] BSTR bstrSectionName, [in] BSTR bstrStringName, [out, retval] BSTR* pVal);

    //显示HTML对话框
    [id(66), helpstring("method ShowHtmlDialog")] HRESULT ShowHtmlDialog([in] BSTR bstrTitle, [in] BSTR bstrURL, [in] LONG nWidth, [in] LONG nHeight, [in] BSTR bstrExtOptions, [in] VARIANT vParam, [out,retval] VARIANT* pvRet);

    //关闭HTML对话框
    [id(67), helpstring("method CloseHtmlDialog")] HRESULT CloseHtmlDialog([in] IDispatch* pdispHtmlDialogDocument, [in] VARIANT vRet);

    //显示一个消息，类似javascript里面的alert
    [id(68), helpstring("method ShowMessage")] HRESULT ShowMessage([in] BSTR bstrText, [in] BSTR bstrTitle, [in] LONG nType, [out,retval] LONG* pnRet);

    //合并URL
    [id(69), helpstring("method CombineURL")] HRESULT CombineURL([in] BSTR bstrBase, [in] BSTR bstrPart, [out,retval] BSTR* pbstrResult);

    //给定URL下载为一个文件
    [id(70), helpstring("method URLDownloadToTempFile")] HRESULT URLDownloadToTempFile([in] BSTR bstrURL, [out,retval] BSTR* pbstrTempFileName);

    //编辑笔记
    [id(71), helpstring("method EditDocument")] HRESULT EditDocument([in] IDispatch* pApp, [in] IDispatch* pEvents, [in] IDispatch* pDocumentDisp, [in] BSTR bstrOptions, [out,retval] LONGLONG* pnWindowHandle);

    //新建一个笔记
    [id(72), helpstring("method NewDocument")] HRESULT NewDocument([in] IDispatch* pApp, [in] IDispatch* pEvents, [in] IDispatch* pFolderDisp, [in] BSTR bstrOptions, [out,retval] LONGLONG* pnWindowHandle);

    //清除所有的笔记窗口
    [id(73), helpstring("method ClearDocumentWindow")] HRESULT ClearDocumentWindow();

    //创建帐户对话框
    [id(74), helpstring("method CreateAccountDlg")] HRESULT CreateAccountDlg([in] BSTR bstrOptions, [out,retval] BSTR* pbstrRetDatabasePath);

    //添加已经存在的帐户对话框
    [id(75), helpstring("method AddExistingAccountDlg")] HRESULT AddExistingAccountDlg([in] BSTR bstrOptions, [out,retval] BSTR* pbstrRetDatabasePath);

    //更改密码对话框
    [id(76), helpstring("method ChangePasswordDlg")] HRESULT ChangePasswordDlg([in] IDispatch* pDatabaseDisp);

    ///保护帐户对话框，已经失效
    [id(77), helpstring("method ProtectAccountDlg")] HRESULT ProtectAccountDlg([in] IDispatch* pDatabaseDisp);

    //更新笔记索引，已经失效
    [id(78), helpstring("method RefreshDatabaseIndex")] HRESULT RefreshDatabaseIndex([in] LONGLONG nWindowHandle, [in] IDispatch* pDatabaseDisp);

    //分享笔记
    [id(79), helpstring("method ShareDocument")] HRESULT ShareDocument([in] LONGLONG nWindowHandle, [in] IDispatch* pDocumentDisp, [in] BSTR bstrParams);

    //分享笔记对话框
    [id(80), helpstring("method ShareDocumentDlg")] HRESULT ShareDocumentDlg([in] IDispatch* pDocumentDisp, [in] BSTR bstrParams);

    //刷新帐户信息
    [id(81), helpstring("method RefreshAccountInfo")] HRESULT RefreshAccountInfo([in] IDispatch* pDatabaseDisp);

    //判断文件是否存在
    [id(82), helpstring("method PathFileExists")] HRESULT PathFileExists([in] BSTR bstrPath, [out,retval] VARIANT_BOOL* pvbExists);

    //从一个文加路径中，获得文件路径，包含最后一个反斜杠\
    [id(83), helpstring("method ExtractFilePath")] HRESULT ExtractFilePath([in] BSTR bstrFileName, [out,retval] BSTR* pbstrFilePath);

    //从一个文加路径中，获得文件文件名，不包含路径
    [id(84), helpstring("method ExtractFileName")] HRESULT ExtractFileName([in] BSTR bstrFileName, [out,retval] BSTR* pbstrName);

    //从一个文件路径中，获得文件扩展名
    [id(85), helpstring("method ExtractFileExt")] HRESULT ExtractFileExt([in] BSTR bstrFileName, [out,retval] BSTR* pbstrExt);

    //从一个文件路径中，获得文件名，并且不包含文件扩展名
    [id(86), helpstring("method ExtractFileTitle")] HRESULT ExtractFileTitle([in] BSTR bstrFileName, [out,retval] BSTR* pbstrTitle);

    //创建目录
    [id(88), helpstring("method CreateDirectory")] HRESULT CreateDirectory([in] BSTR bstrPath);

    //复制文件
    [id(89), helpstring("method CopyFile")] HRESULT CopyFile([in] BSTR bstrExistingFile, [in] BSTR bstrNewFileName);

    //枚举文件夹下面的文件
    [id(90), helpstring("method EnumFiles2")] HRESULT EnumFiles2([in] BSTR bstrPath, [in] BSTR bstrFileExt, [in] VARIANT_BOOL vbIncludeSubFolders, [out,retval] BSTR* pvFiles);

    //枚举文件夹下面的子文件夹
    [id(91), helpstring("method EnumFolders2")] HRESULT EnumFolders2([in] BSTR bstrPath, [out,retval] BSTR* pvFolders);

    //枚举注册表的值
    [id(92), helpstring("method EnumRegValues2")] HRESULT EnumRegValues2([in] BSTR bstrRoot, [in] BSTR bstrKey, [out, retval] BSTR* pvValueNames);

    //枚举注册表的key
    [id(93), helpstring("method EnumRegKeys2")] HRESULT EnumRegKeys2([in] BSTR bstrRoot, [in] BSTR bstrKey, [out, retval] BSTR* pvKeys);

    //从HTML里面提取指定的tag
    [id(94), helpstring("method HtmlExtractTags2")] HRESULT HtmlExtractTags2([in] BSTR bstrHtmlText, [in] BSTR bstrTagName, [in] BSTR bstrTagAttributeName, [in] BSTR bstrTagAttributeValue, [out, retval] BSTR* pvTags);

    //从HTML里面提取连接
    [id(95), helpstring("method HtmlEnumLinks2")] HRESULT HtmlEnumLinks2([in] BSTR bstrHtmlText, [in] BSTR bstrURL, [out, retval] BSTR* pvLinks);

    //获得所有的字体名称
    [id(96), helpstring("method GetFontNames2")] HRESULT GetFontNames2([out,retval] BSTR* pvNames);

    //获得一个路径的长文件名
    [id(98), helpstring("method GetLongPathName")] HRESULT GetLongPathName([in] BSTR bstrPath, [out, retval] BSTR* pbstrRet);

    //获得一个路径的短文件名
    [id(99), helpstring("method GetLongPathName")] HRESULT GetShortPathName([in] BSTR bstrPath, [out, retval] BSTR* pbstrRet);

    //选择一个Wiz文件夹
    [id(100), helpstring("method SelectWizFolder2")] HRESULT SelectWizFolder2([in] BSTR bstrDatabasePath, [in] BSTR bstrPassword, [in] BSTR bstrDescription, [out,retval] IDispatch** ppFolderDisp);

    //通过userid获得对应的数据库路径
    [id(101), helpstring("method UserIdToDatabasePath")] HRESULT UserIdToDatabasePath([in] BSTR bstrUserId, [out,retval] BSTR* ppbstrDatabasePath);

    //显示一个输入文字的对话框
    [id(102), helpstring("method InputBox2")] HRESULT InputBox2([in] OLE_HANDLE hwndParent, [in] VARIANT_BOOL vbPassword, [in] BSTR bstrIconFileName, [in] BSTR bstrTitle, [in] BSTR bstrDescription, [in] BSTR bstrInitValue, [out,retval] BSTR* pbstrText);

    //分享笔记对话框
    [id(103), helpstring("method ShareDocument2")] HRESULT ShareDocument2([in] IDispatch* pApp, [in] IDispatch* pDocumentDisp, [in] BSTR bstrParams);

    //获得文件大小
    [id(104), helpstring("method GetFileSize")] HRESULT GetFileSize([in] BSTR bstrFileName, [out, retval] LONGLONG* pnSize);

    //健一个数字转换为字符串，使用字节，KB，MB等方式
    [id(105), helpstring("method IntToByteSizeString")] HRESULT IntToByteSizeString([in] LONG nSize, [out, retval] BSTR* pbstr);

    //显示帐户信息对话框
    [id(107), helpstring("method AccountInfoDlg2")] HRESULT AccountInfoDlg2([in] IDispatch* pApp, [in] IDispatch* pDatabaseDisp, [in] LONG nFlags);

};
```

在JavaScript里面使用COM安全数组(SafeArray)

JavaScript里面，不能像操作数组那样，直接操作安全数组，必须进行转换。下面是一个例子：

```
var objCommon = new ActiveXObject("WizKMControls.WizCommonUI");
var files = objCommon.EnumFiles("e:\\My Documents\\", "*.txt;*.jpg", true);    //files是一个安全数组
var arrayFile = new VBArray(files).toArray();    //转换为javascript的数组
for (var i = 0; i < arrayFile.length; i++) {
    var optionFile = document.createElement("OPTION");
    optionFile.value = arrayFile[i];
    optionFile.text = arrayFile[i];
    selectFiles.add(optionFile);
}
```
