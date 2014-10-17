WizMeta是WizKMCore.dll包含的一个COM对象。Meta对象必须隶属于一个数据库，因此您不能直接创建这个对象，而是需要通过IWizDatabase对象来获得数据库中的Meta对象。
通过IWizDatabase.Metas，可以获得数据库的所有Meta信息。

Meta，就像ini文件里面的一个记录。Meta信息保存在数据库中，不和服务器同步。利用meta信息，您可以像利用ini文件那样，存储一些用户信息。在Wiz自带的博客发布插件里面，就利用了Meta对象来存储用户数据。

```
 var objApp = new ActiveXObject("WizExplorer.WizExplorerApp");
    var database = objApp.Database;
    var settings_meta = "PublishHtml";
    //
    ...
        settings_meta = "PublishHtml_" + selected_documents.Item(0).Parent.RootFolder.Name;

        //从数据库获得用户设置信息，来进行初始化
        textAddress.value = database.Meta(settings_meta, "Address");
        textPort.value = database.Meta(settings_meta, "Port");
        textUserName.value = database.Meta(settings_meta, "UserName");
        textPassword.value = database.Meta(settings_meta, "Password");
        textImageWidth.value = database.Meta(settings_meta, "ImageWidth");
        checkAutoAddMoreTag.checked = database.Meta(settings_meta, "AutoAddMoreTag") == "1";
        textCategory.value = database.Meta(settings_meta, "Category");
        //

    ...
        //将用户数据保存到数据库里面
        database.Meta(settings_meta, "Address") = address;
        database.Meta(settings_meta, "Port") = port;
        database.Meta(settings_meta, "UserName") = username;
        database.Meta(settings_meta, "Password") = password;
        database.Meta(settings_meta, "ImageWidth") = textImageWidth.value;
        database.Meta(settings_meta, "AutoAddMoreTag") = checkAutoAddMoreTag.checked ? "1" : "0";
        database.Meta(settings_meta, "Category") = textCategory.value;
```

和IWizSettings不同，Meta是保存在数据库中，和数据库相关，而不是全局的设置。

文件：WizKMCore.dll

```
[
    object,
    uuid(8FEE8441-A050-4CF7-99B6-4821DB945D80),
    dual,
    nonextensible,
    helpstring("IWizMeta Interface"),
    pointer_default(unique)
]
interface IWizMeta : IDispatch{


    //获得meta的名称，类似于ini文件里面的Section名称
    [propget, id(1), helpstring("property Name")] HRESULT Name([out, retval] BSTR* pVal);


    //获得meta的key，类似于ini文件里面的key
    [propget, id(2), helpstring("property Key")] HRESULT Key([out, retval] BSTR* pVal);


    //获得/设置meta的值
    [propget, id(3), helpstring("property Value")] HRESULT Value([out, retval] BSTR* pVal);
    [propput, id(3), helpstring("property Value")] HRESULT Value([in] BSTR newVal);


    //获得meta的修改时间
    [propget, id(4), helpstring("property DateModified")] HRESULT DateModified([out, retval] DATE* pVal);


    //删除meta
    [id(5), helpstring("method Delete")] HRESULT Delete(void);
};
```
