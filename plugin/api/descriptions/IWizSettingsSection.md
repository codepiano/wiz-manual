文件：WizKMCore.dll

```
[
    object,
    uuid(E70062FF-7718-4508-8545-3D38A13F3E9C),
    dual,
    nonextensible,
    helpstring("IWizSettingsSection Interface"),
    pointer_default(unique)
]
interface IWizSettingsSection : IDispatch{

    //获得字符串值
    [propget, id(1), helpstring("property StringValue")] HRESULT StringValue([in] BSTR bstrKey, [out, retval] BSTR* pVal);

    //获得整数值
    [propget, id(2), helpstring("property IntValue")] HRESULT IntValue([in] BSTR bstrKey, [out, retval] LONG* pVal);

    //获得布尔值
    [propget, id(3), helpstring("property BoolValue")] HRESULT BoolValue([in] BSTR bstrKey, [out, retval] VARIANT_BOOL* pVal);
};
```
