在这篇教程里面，我们将会制作一个全局插件。

### 了解什么是为知笔记的全局插件(Global)

为知笔记提供了一种在程序运行期间，一直可以运行并且响应Wiz消息的一种插件。例如在wiz运行期间，设定一个定时器，每隔半个小时就可以提醒用户休息眼睛。

### 了解为知笔记插件实现原理


为知笔记运行的时候，就会开启一个脚本运行环境，并进行初始化。在初始化之后，为知笔记会把所有注册成全局脚本的插件对应的脚本，在这个脚本运行环境中运行，并且在为知笔记运行过程中，该脚本运行环境一直有效。

就像一个网页，里面内置了很多脚本，在网页打开的状态下，这些脚本一直都是可以运行的。

而对于直接执行脚本的插件(ExecuteScript)或者html对话框插件(HtmlDialog)，他们只有在用户点击菜单，调用插件的时候，才会被加载执行，当脚本执行完毕，或者用户关闭对话框的时候，插件将会被清除出内存。

### 修改plugin.ini，增加一个简单的全局插件

```
[Common]
FolderName=Hello.World
AppName=HelloWorldApp
AppName_2052=世界你好
AppName_1028=世界你好
AppGUID={edb64fbd-2255-408f-b690-f61e56cb9606}
AppType=Tools
AppVersion=1.0
PluginCount=3


[Plugin_0]
MenuType=Tools
Caption=Hello World
Caption_2052=世界你好
Caption_1028=世界你好
GUID={af91c3db-dc73-4538-b934-b591a6a69aff}
ScriptFileName=HelloWorld.js
Type=ExecuteScript


[Plugin_1]
MenuType=Tools
Caption=Hello World (Html Dialog)
Caption_2052=世界你好 (Html对话框)
Caption_1028=世界你好 (Html对话框)
GUID={6e073661-2a1a-4af0-ad56-abbc18c45666}
Type=HtmlDialog
HtmlDialogFileName=HelloWorld.htm
HtmlDialogTitle=Hello World
HtmlDialogTitle_2052=世界你好
HtmlDialogTitle_1028=世界你好
HtmlDialogWidth=400
HtmlDialogHeight=300


[Plugin_2]
Caption=Hello World (Global)
Caption_2052=世界你好 (全局)
Caption_1028=世界你好 (全局)
GUID={aac081a1-381f-4adc-a7ab-49918de12d97}
ScriptFileName=HelloWorldGlobal.js
Type=Global


[Strings]
labelHelloWorld_2052=世界你好
buttonOK_2052=确定
buttonCancel_2052=取消
OK clicked_2052=点击了确定
Cancel clicked_2052=点击了取消
```

在这里我们增加了一个全局插件功能。

### 编写全局插件脚本

在插件目录新建一个js文件，命名为 HelloWorldGlobal.js，内容如下：

```
function onTimer() {
    objApp.Window.ShowMessage("休息一下眼睛吧", "我的第一个为知笔记插件", 0x40);
}

objApp.Window.AddTimer("onTimer", 1000 * 60 * 30);        //每隔30分钟提醒一次
```

现在，重新启动为知笔记，在程序运行期间，每隔30分钟，就会出现以下消息框，提示你休息一下眼睛。

### 给全局插件进行国际化

在上面的消息提示中，我们直接采用了中文。记下来我们要给他实现国际化和本地化。

首先，全局插件无法直接获得插件所在的路径，这是因为所有的全局插件，都是在一个插件运行环境里面执行的。因此我们需要通过其他手段来获得插件安装路径，然后从插件安装路径的plugin.ini文件里面获得本地化的字符串。

1. 获得插件安装路径
    IWizExplorerApp 提供了一个方法：GetPluginPathByScriptFileName，可以通过一个脚本的文件名，获得相应的路径。

    因此，我们可以使用下面的代码：

    ```
    var pluginPath = objApp.GetPluginPathByScriptFileName("HelloWorldGlobal.js");
    ```

    获得插件路径。这个路径包含插件安装完整路径，包含最后面的反斜杠（\）

1. 从文件中获得本地化的字符串
    同样，IWizExplorerApp提供了 LoadStringFromFile方法，来从一个ini文件里面获得本地化的字符串。

1. 修改脚本文件
    ```
    var pluginPath = objApp.GetPluginPathByScriptFileName("HelloWorldGlobal.js");
    var pluginFileName = pluginPath + "plugin.ini";

    function getLocalizedString(str) {
        return objApp.LoadStringFromFile(pluginFileName, str);
    }

    function onTimer() {
        objApp.Window.ShowMessage(getLocalizedString("Please take a break"), getLocalizedString("My first plugin"), 0x40);
    }

    objApp.Window.AddTimer("onTimer", 1000 * 60 * 30);        //每隔30分钟提醒一次
    ```

    在脚本中，我们首先获得了插件路径以及plugin.ini的文件路径名，然后编写了getLocalizedString方法。

    然后在消息提示框里面，我们使用本地化的语言进行显示。

1. 修改plugin.ini，增加本地化语言
    ```
    [Common]
    FolderName=Hello.World
    AppName=HelloWorldApp
    AppName_2052=世界你好
    AppName_1028=世界你好
    AppGUID={edb64fbd-2255-408f-b690-f61e56cb9606}
    AppType=Tools
    AppVersion=1.0
    PluginCount=3


    [Plugin_0]
    MenuType=Tools
    Caption=Hello World
    Caption_2052=世界你好
    Caption_1028=世界你好
    GUID={af91c3db-dc73-4538-b934-b591a6a69aff}
    ScriptFileName=HelloWorld.js
    Type=ExecuteScript


    [Plugin_1]
    MenuType=Tools
    Caption=Hello World (Html Dialog)
    Caption_2052=世界你好 (Html对话框)
    Caption_1028=世界你好 (Html对话框)
    GUID={6e073661-2a1a-4af0-ad56-abbc18c45666}
    Type=HtmlDialog
    HtmlDialogFileName=HelloWorld.htm
    HtmlDialogTitle=Hello World
    HtmlDialogTitle_2052=世界你好
    HtmlDialogTitle_1028=世界你好
    HtmlDialogWidth=400
    HtmlDialogHeight=300


    [Plugin_2]
    Caption=Hello World (Global)
    Caption_2052=世界你好 (全局)
    Caption_1028=世界你好 (全局)
    GUID={aac081a1-381f-4adc-a7ab-49918de12d97}
    ScriptFileName=HelloWorldGlobal.js
    Type=Global


    [Strings]
    labelHelloWorld_2052=世界你好
    buttonOK_2052=确定
    buttonCancel_2052=取消
    OK clicked_2052=点击了确定
    Cancel clicked_2052=点击了取消


    Please take a break_2052=请休息一下
    My first plugin_2052=我的第一个插件
    ```

### 避免全局脚本命名冲突

前面提到了，所有的全局插件都是在一个插件运行环境下执行的，因此，如果全局插件比较多，那么就可能会出现插件变量或者函数的名称，产生冲突的问题，例如前面的pluginPath, pluginFileName, onTimer，都不是一个合适的命名方式，因为其它插件，也可能会使用相同的名称或者变量，导致运行的时候，出现不可预期的结果。

为了避免这个问题，我们需要给自己的变量和函数名，加上前缀或者后缀，例如我们给所有的变量增加一个前缀，叫做HelloWorld，这样就可以有效避免这个问题了。

修改后HelloWorldGlobal.js的代码如下：

```
var HelloWorld_pluginPath = objApp.GetPluginPathByScriptFileName("HelloWorldGlobal.js");
var HelloWorld_pluginFileName = HelloWorld_pluginPath + "plugin.ini";

function HelloWorld_getLocalizedString(str) {
    return objApp.LoadStringFromFile(HelloWorld_pluginFileName, str);
}

function HelloWorld_onTimer() {
    objApp.Window.ShowMessage(HelloWorld_getLocalizedString("Please take a break"), HelloWorld_getLocalizedString("My first plugin"), 0x40);
}

objApp.Window.AddTimer("HelloWorld_onTimer", 1000 * 10);        //每隔30分钟提醒一次
```

修改后的HelloWorldGlobal.js，就不用担心命名冲突的问题了。
