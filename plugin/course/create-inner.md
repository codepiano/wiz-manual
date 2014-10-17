为知笔记使用了IE作为脚本运行环境，因此可以使用某些IE提供的功能，例如创建一个COM对象的时候，可以用 new ActiveXObject("xxx.xxxx")。

为知笔记提供的内部对象，也都是基于COM的，但是开发人员不应该使用 new ActiveXObject 这种方式来创建为知笔记内部的COM对象，因为 new ActiveXObject 依赖于注册表创建COM对象，要求COM对象必须进行注册。但是为知笔记是可以按照绿色方式运行的，所有的COM组件都不需要注册，因此，必须采用为知笔记自己提供的方式来创建对象。

1. 获得 IWizExplorerApp (objApp)对象，请参考：

    [内部对象](./course/inner-object.md)

    该文章内，详细介绍了如何在插件里面获得IWizExplorerApp对象(objApp)。

1. 使用objApp.CreateWizObject来创建为知笔记对象。

    下面就是一些例子：

    ```
    var objCommon = objApp.CreateWizObject("WizKMControls.WizCommonUI");
    var objDestDatabase = objApp.CreateWizObject("WizKMCore.WizDatabase");
    var objProgress = objApp.CreateWizObject("WizKMControls.WizProgressWindow");
    ```

    其中CreateWizObject的参数是ProgID，类似于new ActiveXObject的参数。详细的对象列表，请参看：

    为知笔记内部对象索引

1. 使用objApp.CreateActiveXObject来创建外部COM对象。

    因为IE的安全等级限制，某些COM对象无法在Wiz插件里面使用，因此，可以用过Wiz提供的功能，来避免这个问题。就是用objApp.CreateActiveXObject 来代替 new ActiveXObject。例如下面的代码：

    ```
    //避免ActiveX对象安全警告
    var spVoice = objApp.CreateActiveXObject("SAPI.SpVoice");
    ```
