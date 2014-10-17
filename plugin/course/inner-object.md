接下来的教程，我们将会深入为知笔记内部的对象，来探索为知笔记内部对象的奥秘以及组织方式。

### 内部对象IWizExplorerApp

IWizExplorerApp是为知笔记主程序运行的时候，对外暴露的一个内部对象，通过IWizExplorerApp，可以获得为知笔记正在打开的账户数据，主窗口各种控件等等。在三种插件里面，都提供了脚本直接访问这个对象的方法。

### 插件脚本获得为知笔记内部对象的方法

#### ExecuteScript类型的插件：

这种插件的脚本，在运行的时候，内置了一个对象：WizExplorerApp。他的类型就是IWizExplorerApp。

在 [新建和发布插件](./course/hello-world.html) 里面，我们就使用了这个对象：

```
var objApp = WizExplorerApp;
var objWindow = objApp.Window;
objWindow.ShowMessage("世界你好", "我的第一个为知笔记插件", 0x40);
```

如图，我们定一个objApp，直接引用了 内置对象 WizExplorerApp。

#### HtmlDialog类型的插件：

在 [弹出对话框](./course/create-dialog.html) 里面，我们在html里面也获得了这个对象：

var objApp = window.external;

window.external，就是内部浏览器保留的一个对象，用于获得IWizExplorerApp 。

#### Global类型的插件：

为知笔记在运行插件脚本之前，也内置了一个对象：WizExplorerApp，类型同样是IWizExplorerApp。另外，Global类型插件，在运行的时候，也与定义了objApp对象，同样是IWizExplorerApp。

在 [全局插件](./course/global-plugin.html) 里面，我们直接使用了objApp，而没有进行定义，是因为全局插件已经帮我们定义好了。

```
function onTimer() {
    objApp.Window.ShowMessage("休息一下眼睛吧", "我的第一个为知笔记插件", 0x40);
}

objApp.Window.AddTimer("onTimer", 1000 * 60 * 30);        //每隔30分钟提醒一次
```

为了统一，我们在以后的教程中，凡是提到objApp的，就代表IWizExplorerApp对象，并且我们已经通过某种方式获得了这个对象。

### 通过IWizExplorerApp来获得其他的对象

这个对象详细的文档，请参考接口描述文件：IWizExplorerApp

获得当前打开的账户数据库(IWizDatabase)

通过这个对象，可以获得用户的账户数据。

```
var objDatabase = objApp.Database;
```

注意：在全局插件中，已经直接定义了 objDatabase对象，在插件脚本中可以直接使用。

获得为知笔记主窗口对象(IWizExplorerWindow)

通过这个对象，可以获得主窗口的一些控件，在全局插件中，还可以响应主窗口的一些事件。

```
var objWindow = objApp.Window;
```

注意：在全局插件中，一定直接定义了objWindow对象，在插件脚本中可以直接使用。

一些常用的功能的对象(IWizCommonUI)

这个对象提供了大量辅助性的功能，例如ini文件读写，注册表文件读写，文本文件读写等等。

```
var objCommon = objApp.CreateWizObject("WizKMControls.WizCommonUI");
```

注意：在全局插件中，一定直接创建了objCommon 对象，在插件脚本中可以直接使用。

关于各种对象的描述，请参看接口描述部分。
