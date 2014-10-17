## Hello World

在[新建和发布插件](./course/hello-world.html)中，我们制作了第一个插件，接下来我们将会详细介绍插件里面的两个文件：

### plugin.ini

plugin.ini是一个典型的ini文件，里面有两个section，分别是[Common] 和 [Plugin_0]。

```
[Common]
FolderName=Hello.World
AppName=HelloWorldApp
AppName_2052=世界你好
AppName_1028=世界你好
AppGUID={edb64fbd-2255-408f-b690-f61e56cb9606}
AppType=Tools
AppVersion=1.0
PluginCount=1

[Plugin_0]
MenuType=Tools
Caption=Hello World
Caption_2052=世界你好
Caption_1028=世界你好
GUID={af91c3db-dc73-4538-b934-b591a6a69aff}
ScriptFileName=HelloWorld.js
Type=ExecuteScript
```

其中[Common]部分，是告诉Wiz，这个插件的一些信息，例如：

+ FolderName：插件安装后的文件夹的名称，给插件一个让人容易理解的名称。如果没有，那么就会使用插件的AppGUID作为文件夹名称
+ AppName：插件的名称，这个名称，会出现在插件管理器里面。这个字符串可以进行本地化，例如简体中文本地化名称：AppName_2052，繁体中文 AppName_1028。后面的数字代表某一个区域或者语言，可以参考：http://msdn.microsoft.com/zh-cn/subscriptions/downloads/0h88fahh(v=vs.80).aspx
+ AppGUID：插件的唯一标识，每一个插件都不应该相同，可以到这里直接生成：
+ http://www.guidgen.com/
+ AppType：表明插件的类型，可以是任意字符串
+ AppVersion：插件版本，应该是一个小数，通常是1.0之类的
+ PluginCount：插件功能的数量，例如在这个例子中，我们只有一个插件功能
+ MenuType：插件菜单的类型，这个菜单应该先是在哪一个菜单里面。
+ Caption：菜单文字，可以进行本地化，类似[Common]里面的AppName。
+ GUID：插件功能的GUID，每一个插件功能的GUID都不应该相同，可以像[Common]里面的AppGUID那样，在线生成一个。
+ ScriptFileName：当用户点击插件菜单后，执行的脚本文件。例如我们指定了 HelloWorld.js文件
+ Type：插件功能类型，我们在这里指定了是执行一个脚本文件，也就是ScriptFileName指定的文件。

Plugin_x，则代表某一个插件的描述部分，x从0开始，到插件数量-1。如果有多个插件功能，每一个插件功能有一个独立的Plugin_x section，例如[Plugin_0], [Plugin_1] 等等

其中[Plugin_x]中的MenuType，有以下几种类型：

+ DocumentCtrlAdvanced            文档右键菜单，高级下面，用于一些不常用的高级功能
+ DocumentCtrl                    文档右键菜单
+ CategoryCtrlAllFolders          左边的目录树，出现在所有文件夹右键菜单
+ CategoryCtrlAllTags             目录树，出现在所有标签右键菜单
+ CategoryCtrlAllStyles           目录树，出现在所有样式右键菜单
+ CategoryCtrlAllQuickSearches    目录树，出现在所有快速搜索右键菜单
+ CategoryCtrlFolders             目录树，出现在用户右键点击某一个文件夹右键菜单
+ CategoryCtrlDocumentsAdvanced   目录树，出现在用户右键点击某一个文档右键菜单（高级）
+ CategoryCtrlDocuments           目录树，出现在用户右键点击某一个文档右键菜单
+ CategoryCtrlTags                目录树，出现在用户右键点击某一个标签右键菜单
+ CategoryCtrlStyles              目录树，出现在用户右键点击某一个样式右键菜单
+ CategoryCtrlQuickSearches       目录树，出现在用户右键点击某一个快速搜索右键菜单
+ AttachmentCtrl                  附件列表，出现在用户右键点击附件列表
+ Import                          导入功能，例如导入文件，导入博客
+ Export                          导出功能，例如导出文件
+ Tools                           工具，出现在工具菜单下面
+ NoteTools                       文档工具，出现在文档阅读/编辑右上方更多按钮下面
+ NoteShare                       文档分享，出现在文档阅读/编辑右上方分享按钮下面

如果没有指定MenuType，那么插件菜单将会显示在Tools菜单下面，也就是相当于Tools。

其中[Plugin_x]中的Type，有以下几种类型：

+ HtmlDialog：当用户点击菜单后，显示一个html对话框，对话框的URL可以是一个本地文件，也可以是一个远程的URL
+ ExecuteScript：当用户点击菜单后，执行一个脚本。我们的HelloWorld就是这种类型
+ Global：不添加任何菜单，但是在Wiz启动的时候，就执行执行的脚本文件，并且在Wiz运行期间，一直保持活动状态
+ HtmlSaver：不添加任何菜单，而是在保存网页的时候，执行指定的脚本文件，处理保存后的网页。通常用户在保存网页之后，处理保存后的内容。

对于Type，除了HtmlDialog类型外，其他的类型，都必须指定ScriptFileName，来指定脚本文件（必须是js文件）。

对于HtmlDialog，我们将会在 [弹出对话框](./course/create-dialog.html) 里面详细说明。

### HelloWorld.js

helloworld.js文件，就是在用户点击菜单后，执行的脚本文件。这个文件很简单，只有如下三行：

```
var objApp = WizExplorerApp;
var objWindow = objApp.Window;
objWindow.ShowMessage("世界你好", "我的第一个为知笔记插件", 0x40);
```

WizExplorerApp是插件脚本内置的一个对象，类型是 IWizExplorerApp。我们定义了一个objApp，是为了更边方便输入。

然后我们获得Wiz窗口对象，并且显示了一个消息框。
