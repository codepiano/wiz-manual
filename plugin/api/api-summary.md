为知笔记Windows客户端开放了大量的API，其中绝大部分，都通过COM提供，可以在javascript, C#, C++, Delphi等语言中使用。

接口通过IDL(Interface description language)语言描述。

IDL语言，可以参考：
[Interface_description_language · Wiki](http://en.wikipedia.org/wiki/Interface_description_language)

[IDL_百度百科](http://baike.baidu.com/view/160644.htm)

1. [WizKMCore.idl](/manual/plugin/api/definitions/WizKMCore.html) 描述了为知笔记内部对象的接口
1. [WizKMControls.idl](/manual/plugin/api/definitions/WizKMControls.html) 描述了为知笔记主要的界面控件接口
    1. [IWizDocumentListCtrl](/manual/plugin/api/descriptions/IWizDocumentListCtrl.html)
    1. [IWizDocumentAttachmentListCtrl](/manual/plugin/api/descriptions/IWizDocumentAttachmentListCtrl.html)
    1. [IWizCommonUI](/manual/plugin/api/descriptions/IWizCommonUI.html)
    1. [IWizBatchDownloader](/manual/plugin/api/descriptions/IWizBatchDownloader.html)
    1. [IWizCategoryCtrl](/manual/plugin/api/descriptions/IWizCategoryCtrl.html)
1. [Wiz.idl](/manual/plugin/api/definitions/Wiz.html) 描述了为知笔记主程序的一些接口
    1. [IWizExplorerApp](/manual/plugin/api/descriptions/IWizExplorerApp.html)
    1. [IWizExplorerWindow](/manual/plugin/api/descriptions/IWizExplorerWindow.html)
1. [WizTools.idl](/manual/plugin/api/definitions/WizTools.html) 描述了为知笔记一些通用功能的接口
1. [WizWebCapture.idl](/manual/plugin/api/definitions/WizWebCapture.html) 描述了为知笔记网页捕捉功能的接口
1. [WizWebCaptureResourceToDocument.idl](/manual/plugin/api/definitions/WizWebCaptureResourceToDocument.html) 描述了为知笔记网页捕捉的一些接口

适用于为知笔记3.0.22以及以上版本
