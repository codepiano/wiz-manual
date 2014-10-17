文件：WizKMCore.dll

```
[
object,
uuid(74D15AA2-9E93-437C-990B-81A3D9155C06),
dual,
nonextensible,
helpstring("IWizRowset Interface"),
pointer_default(unique)
]
interface IWizRowset : IDispatch{

//获得列数
[propget, id(1), helpstring("property ColumnCount")] HRESULT ColumnCount([out, retval] LONG* pVal);

//判断是否结束
[propget, id(2), helpstring("property EOF")] HRESULT EOF([out, retval] VARIANT_BOOL* pVal);

//获得字段的值。转换为字符串
[id(3), helpstring("method GetFieldValue")] HRESULT GetFieldValue([in] LONG nFieldIndex, [out,retval] BSTR* pbstrFieldValue);

//获得字段值，作为variant
[id(4), helpstring("method GetFieldValue2")] HRESULT GetFieldValue2([in] LONG nFieldIndex, [out,retval] VARIANT* pvRet);

//获得字段名称
[id(5), helpstring("method GetFieldName")] HRESULT GetFieldName([in] LONG nFieldIndex, [out,retval] BSTR* pbstrFieldName);

//获得字段类型
[id(6), helpstring("method GetFieldType")] HRESULT GetFieldType([in] LONG nFieldIndex, [out,retval] BSTR* pbstrFieldType);

//移动到下一行
[id(7), helpstring("method MoveNext")] HRESULT MoveNext(void);

//从一个字段中读取二进制流数据。其中vstream必须实现了IStream接口，在脚本中，通常可以使用ADODB.Stream这个COM对象。
[id(8), helpstring("method GetFieldValueAsStream")] HRESULT GetFieldValueAsStream([in] LONG nFieldIndex, [in] VARIANT vStream);
};
```
