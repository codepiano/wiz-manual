在本次教程内，我们将会开发一个真实的插件：MathJax，可以用MathJax显示笔记中的公式。

### 关于MathJax

MathJax 是一个开源的基于 Ajax 的数学公式显示的解决方案，结合多种先进的Web技术，支持主流的浏览器。MathJax 根据页面中定义的 LaTex 数据，生成对应的数学公式。

### 编写思路

首先，如果一个笔记中包含LaTeX 或者 MathML 编写的公式，那么在浏览的时候，我们可以动态加入MathJax渲染引擎，这样就可以将渲染笔记里面的公式了。

在上一篇教程内，我们了解了如何响应为知笔记打开笔记的消息，因此我们将会在笔记HTML完成的时候，来判断当前笔记是否需要使用MatjJax来渲染，如果需要，则动态插入一行脚本，来渲染公式。

### plugin.ini

```
[Common]
FolderName=Notes.MathJax
AppName=MathJax
AppName_2052=MathJax
AppName_1028=MathJax
AppGUID={11e7ec35-a748-415f-bd1e-a39a95adb873}
AppType=Display
AppVersion=1.0
PluginCount=2


[Plugin_0]
MenuType=Editor
Caption=Using MathJax
Caption_2052=使用MathJax语法渲染
Caption_1028=使用MathJax语法渲染
GUID={70d7198a-4184-4459-93cd-c53ef29d7da8}
ScriptFileName=MarkAsMathjax.js
Type=ExecuteScript


[Plugin_1]
Caption=MathJax
Caption_2052=MathJax
Caption_1028=MathJax
GUID={c49879c2-b29e-45d3-8f0b-852bf1bdbd93}
ScriptFileName=MathJaxGlobal.js
Type=Global
```

我们定义了两个插件功能，

第一个会在编辑菜单右边的箭头下面添加一个菜单：使用MathJax语法渲染。当用户点击这个菜单后，将会标记这个笔记，需要使用MathJax来渲染。

第二个插件功能，是一个全局插件，我们需要监听为知笔记打开笔记的消息，在合适的时候，使用MathJax来渲染公式。

### 标记笔记需要使用MathJax来渲染

MarkAsMathjax.js代码如下：

```
var objApp = WizExplorerApp;
var objWindow = objApp.Window;
var objDocument = objWindow.CurrentDocument;
if (objDocument) {
    objDocument.Type = "Mathjax";
    //
    objApp.AddGlobalScript(objApp.CurPluginAppPath + "MathjaxCurrentDocument.js");
}
```

当用户点击菜单后，获得为知笔记当前打开的笔记对象，然后把它的类型修改成 Mathjax。

这样当在为知笔记里面，打开这个笔记的时候，就不需要再次进行标记，插件可以自动的判断出来是否需要使用MathJax渲染。

接下来的代码：

```
objApp.AddGlobalScript(objApp.CurPluginAppPath + "MathjaxCurrentDocument.js");
```

objApp.CurPluginAppPath，返回了当前插件的安装路径。注意，这个属性，不能用于全局插件，只能用于ExecuteScript或者HtmlDialog类型的插件。

objApp.AddGlobalScript()，则向全局插件添加了脚本用于执行。在上一教程内，我们讲过，所有的全局插件都在一个插件脚本运行环境内运行，包括我们这个插件应用里面的第二个插件功能（全局插件）的脚本(MathJaxGlobal.js)。同时，通过objApp.AddGlobalScript，可以动态添加一个js文件到全局插件运行环境进行运行。

下面我们看看MathjaxCurrentDocument.js的文件内容：

```
mj_addMathjaxScriptToCurrentDocument();
```

这个文件只有一行，就是添加MathJax引擎到当前笔记里面。但是这个函数的内容，这个文件里面并没有，而是在全局插件脚本文件(MathJaxGlobal.js)里面。能够正常运行，是因为全局插件脚本和通过obj.AddGlobalScript动态添加的脚本，是在一起运行的。

### 监听为知笔记消息，渲染包含公式的笔记

我们来看看全局插件的脚本(MathjaxGlobal.js)：

```
function mj_addMathjaxScript(doc) {
    if (!doc)
        return;

    var elem = doc.createElement("script");
    elem.src = "http://cdn.mathjax.org/mathjax/latest/MathJax.js?config=TeX-AMS-MML_HTMLorMML";
    doc.body.appendChild(elem);
}

function mj_addMathjaxScriptToCurrentDocument() {
    var doc = objApp.CurrentDocumentHtmlDocument;
    mj_addMathjaxScript(doc);
}

function mj_onHtmlDocumentCompleted(doc) {
    try {
        var objDocument = objApp.Window.CurrentDocument;
        if (objDocument) {
            if (objDocument.Type == "Mathjax") {
                mj_addMathjaxScript(doc);
            }
        }
    }
    catch (err) {
    }
}

eventsHtmlDocumentComplete.add(mj_onHtmlDocumentCompleted);
```

mj_addMathjaxScript里面，创建了一个script对象，并且引用了MathJax的渲染引擎。

mj_addMathjaxScriptToCurrentDocument里面，则是向当前笔记添加。需要注意的是，这个函数，在当前脚本里面没有引用，而是在前面所说的obj.AddGlobalScript里面，动态添加的脚本里面引用了。

接下来，就是响应为知笔记的事件，并且判断当前笔记是否需要进行渲染（通过objDocument.Type）。

如果是，则进行渲染。
