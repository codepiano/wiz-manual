IWizBatchDownloader是WizKMControls.dll包含的一个COM对象。IWizBatchDownloader主要用来进行批量下载网页或者导入文档。

IWizBatchDownloader可以批量下载网页，或者导入文件，并且允许在导入后合并文档。在Wiz自带的博客下载插件以及导入文件插件中，都使用了这个对象，您可以直接查看插件的源代码，获得一些使用的例子。

ProgID：WizKMControls.WizBatchDownloader
文件：WizKMControls.dll

```
[
    object,
    uuid(49C2236A-6DA1-4AA9-BAD7-BF0DA94535A4),
    dual,
    nonextensible,
    helpstring("IWizBatchDownloader Interface"),
    pointer_default(unique)
]
interface IWizBatchDownloader : IDispatch{

    //添加一个下载任务。
    //bstrDatabasePath：数据库路径
    //bstrLocation：文档保存位置
    //bstrURL：需要下载的网页URL或者需要导入的文件名。可以是doc/text/ppt/xls/jpg/bmp/png等等文件
    //bstrLinkText：链接文字，可以为空字符串
    //nUpdateDocumentFlags：更新文档的选项，请参看IWizDocument.UpdateDocument方法
    //vbLinkTextAsTitle：是否将链接文字作为文档标题
    //vbExecuteScript：是否执行html里面的代码
    [id(1), helpstring("method AddJob")] HRESULT AddJob([in] BSTR bstrDatabasePath, [in] BSTR bstrLocation, [in] BSTR bstrURL, [in] BSTR bstrLinkText, [in] LONG nUpdateDocumentFlags, [in] VARIANT_BOOL vbLinkTextAsTitle, [in] VARIANT_BOOL vbExecuteScript);

    //从text文件添加链接。text文件里面，每行一个链接。
    //bstrTextFileName：text文件名。
    //其它参数同AddJob方法
    [id(2), helpstring("method AddJobsFromTextFile")] HRESULT AddJobsFromTextFile([in] BSTR bstrDatabasePath, [in] BSTR bstrLocation, [in] BSTR bstrTextFileName, [in] LONG nUpdateDocumentFlags, [in] VARIANT_BOOL vbExecuteScript);

    //显示批量下载窗口。vbStartDownload：是否自动开始下载
    [id(3), helpstring("method ShowWindow")] HRESULT ShowWindow(VARIANT_BOOL vbStartDownload);

    //获得/设置是否在下载完成后，将下载的文档合并成一个新的文档
    [propget, id(4), helpstring("property CombineDocuments")] HRESULT CombineDocuments([out, retval] VARIANT_BOOL* pVal);
    [propput, id(4), helpstring("property CombineDocuments")] HRESULT CombineDocuments([in] VARIANT_BOOL newVal);

    //添加任务，保留。
    [id(5), helpstring("method AddJob2")] HRESULT AddJob2([in] BSTR bstrDatabasePath, [in] BSTR bstrLocation, [in] BSTR bstrURL, [in] BSTR bstrLinkText, [in] LONG nUpdateDocumentFlags, [in] LONG nJobFlags);

    //添加任务，允许指定覆盖一个文档。

    [id(6), helpstring("method AddJob3")] HRESULT AddJob3([in] BSTR bstrDatabasePath, [in] BSTR bstrOverwriteDocumentGUID, [in] BSTR bstrURL, [in] BSTR bstrTitle, [in] LONG nUpdateDocumentFlags, [in] LONG nJobFlags);

    //获得，更改窗口标题

    [propget, id(7), helpstring("property Title")] HRESULT Title([out, retval] BSTR* pVal);
    [propput, id(7), helpstring("property Title")] HRESULT Title([in] BSTR newVal);

    //保存当前工作状态为文件
    [id(8), helpstring("method SaveJobsToFile")] HRESULT SaveJobsToFile([in] BSTR bstrFileName);


    //从文件里面获得恢复工作列表
    [id(9), helpstring("method LoadJobsFromFile")] HRESULT LoadJobsFromFile([in] BSTR bstrFileName);

    //添加一个任务
    [id(10), helpstring("method AddJob")] HRESULT AddJob4([in] BSTR bstrDatabasePath, [in] BSTR bstrLocation, [in] BSTR bstrURL, [in] BSTR bstrLinkText, [in] LONG nUpdateDocumentFlags, [in] VARIANT_BOOL vbLinkTextAsTitle, [in] VARIANT_BOOL vbExecuteScript, [in] BSTR bstrKbGUID, [in] BSTR bstrTagGUIDs);


};
```

具体使用例子，请参考Wiz自带的博客下载插件和导入文件插件。
