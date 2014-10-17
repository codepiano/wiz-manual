IWizSQLiteDatabase 数据库对象，可以在本地生成一个sqlite数据库，并且可以指定简单的sql语句以及查询

文件：WizKMCore.dll

```
[
    object,
    uuid(AAD77AF9-FDD1-454A-A52F-7E2037EB5F8D),
    dual,
    nonextensible,
    helpstring("IWizSQLiteDatabase Interface"),
    pointer_default(unique)
]
interface IWizSQLiteDatabase : IDispatch{

    //打开一个sqlite数据库
    [id(1), helpstring("method Open")] HRESULT Open([in] BSTR bstrFileName);

    //关闭
    [id(2), helpstring("method Close")] HRESULT Close(void);

    //执行一个查询语句，返回IWizRowset
    [id(3), helpstring("method SQLQuery")] HRESULT SQLQuery([in] BSTR bstrSQL, [in] BSTR bstrOptions, [out,retval] IDispatch** ppRowsetDisp);

    //执行一个update/delete语句，返回受影响的行数
    [id(4), helpstring("method SQLExecute")] HRESULT SQLExecute([in] BSTR bstrSQL, [in] BSTR bstrOptions, [out,retval] LONG* pnRowsAffect);
};
```
