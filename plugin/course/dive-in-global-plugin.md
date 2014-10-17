我们在之前简单制作了一个全局插件，接下来将会深入研究这种类型插件的运行原理以及使用方法。

### 全局插件的运行原理

在为知笔记启动的时候，会创建一个脚本执行环境（类似于chrome的V8引擎），然后执行插件环境初始化。接下来，会把所有的全局脚本放在这个引擎里面执行。同时这个引擎，会一直有效，直到关闭为知笔记。

而ExecuteScript类型的脚本，实际上和全局脚本类似，之不过，这种类型的插件，只有在用户点击菜单的时候，才会生成一个新的脚本执行环境，来执行脚本，并且在脚本执行完毕后，销毁执行环境。同时，每一个插件，执行的时候都是独立的，和其他插件无关。

### 全局插件的初始化

从前面的介绍可以看出，全局插件有一个初始化过程，而这个初始化，实际上就是执行了一段脚本。该脚本保存在为知笔记安装路径下面的 plugins文件夹，文件名是global_plugin_share.js，下面我们打开看看它的内容：

```
/*
////定义一些全局的变量，在所有的脚本插件之前实现////
*/

/*
////WizExplorer内部对象，所有的全局插件脚本都可以使用////
*/
var objApp = WizExplorerApp;
var objDatabase = objApp.Database;
var objWindow = objApp.Window;
var objCommon = objApp.CreateWizObject("WizKMControls.WizCommonUI");

function WizAlert(msg) {
    objWindow.ShowMessage(msg, "Wiz", 0x00000040);
}

function WizConfirm(msg) {
    return objWindow.ShowMessage(msg, "Wiz", 0x00000020 | 0x00000001) == 1;
}


function WizEventDispatcher() {
    this.listeners = [];
    this.add = function(callback) {
        this.listeners.push(callback);
        return callback;
    }
    this.remove = function(callback) {
        for (var i = 0; i < this.listeners.length; i++) {
            if (this.listeners[i] == callback) {
                this.listeners.splice(i, 1);
                return true;
            }
        }
        return false;
    }
    this.dispatch = function() {
        for (var i = 0; i < this.listeners.length; i++) {
            var callback = this.listeners[i];
            try {
                callback();
            }
            catch (err) {
            }
        }
    }
    this.dispatch1 = function(arg1) {
        for (var i = 0; i < this.listeners.length; i++) {
            var callback = this.listeners[i];
            try {
                callback(arg1);
            }
            catch (err) {
            }
        }
    }
    this.dispatch2 = function(arg1, arg2) {
        for (var i = 0; i < this.listeners.length; i++) {
            var callback = this.listeners[i];
            try {
                callback(arg1, arg2);
            }
            catch (err) {
            }
        }
    }
    this.dispatch3 = function(arg1, arg2, arg3) {
        for (var i = 0; i < this.listeners.length; i++) {
            var callback = this.listeners[i];
            try {
                callback(arg1, arg2,arg3);
            }
            catch (err) {
            }
        }
    }
}


var eventsAccountChanged = new WizEventDispatcher();
var eventsClose = new WizEventDispatcher();
var eventsTabCreate = new WizEventDispatcher();
var eventsTabClose = new WizEventDispatcher();
var eventsHtmlDocumentComplete = new WizEventDispatcher();
var eventsDocumentBeforeEdit = new WizEventDispatcher();
var eventsDocumentBeforeChange = new WizEventDispatcher();

function WizOnAccountChnaged() {
    eventsAccountChanged.dispatch();
}


function WizOnClose() {
    eventsClose.dispatch();
}

function WizOnTabCreate(doc) {
    eventsTabCreate.dispatch1(doc);
}

function WizOnTabClose(objHtmlDocument, objWizDocument) {
    eventsTabClose.dispatch2(objHtmlDocument, objWizDocument);
}

function WizOnHtmlDocumentComplete(doc) {
    eventsHtmlDocumentComplete.dispatch1(doc);
}

function WizOnDocumentBeforeEdit(objHtmlDocument, objWizDocument) {
    return eventsDocumentBeforeEdit.dispatch2(objHtmlDocument, objWizDocument);
}
function WizOnDocumentBeforeChange(objHtmlDocument, objWizDocumentOld, objWizDocumentNew) {
    return eventsDocumentBeforeChange.dispatch3(objHtmlDocument, objWizDocumentOld, objWizDocumentNew);
}

function WizComDateTimeToDate(com_dt) {
    var dt = new Date(Date.parse(com_dt));
    return dt;
}

function WizFormatInt2(val) {
    if (val < 10)
        return "0" + val;
    else
        return "" + val;
}

function WizComDateTimeToStr(com_dt) {
    var dt = new Date(Date.parse(com_dt));
    return "" + dt.getFullYear() + "-" + WizFormatInt2(dt.getMonth() + 1) + "-" + WizFormatInt2(dt.getDate()) + " " + WizFormatInt2(dt.getHours()) + ":" + WizFormatInt2(dt.getMinutes()) + ":" + WizFormatInt2(dt.getSeconds());
}


/*
function myonclose() {
    WizAlert("onclose");
}
eventsClose.add(myonclose);

function myontabcreate(doc) {
    WizAlert("ontabcreate");
}
eventsTabCreate.add(myontabcreate);

function myontabclose(doc) {
    WizAlert("ontabclose");
}
eventsTabClose.add(myontabclose);

function myonhtmldocumentcomplete(doc) {
    WizAlert("onhtmldocumentcomplete");
}
eventsHtmlDocumentComplete.add(myonhtmldocumentcomplete);
*/
```

首先是下面的代码：

```
var objApp = WizExplorerApp;
var objDatabase = objApp.Database;
var objWindow = objApp.Window;
var objCommon = objApp.CreateWizObject("WizKMControls.WizCommonUI");
```

在这里我们与定义了几个对象：objApp, objDatabase, objWindow以及objCommon。

这几个对象的用于可以参考相关的说明。这几个对象，可以在所有的全局插件脚本里面直接使用，而不需要重新定义。

```
function WizAlert(msg) {
    objWindow.ShowMessage(msg, "Wiz", 0x00000040);
}

function WizConfirm(msg) {
    return objWindow.ShowMessage(msg, "Wiz", 0x00000020 | 0x00000001) == 1;
}
```

接下来，定义了两个函数，提供了javascript里面的alert, confirm 这两个函数的功能。因为为知笔记插件执行环境，本身并没有提供alert以及prompt这两个函数（他们实际上是html dom里面的功能）。

因此在全局插件脚本里面，可以直接使用 WizAlert来代替alert, WizConfirm来代替confirm。

接下来的脚本，实现了监听为知笔记一些消息的功能。

```
function WizEventDispatcher() {
    this.listeners = [];
    this.add = function(callback) {
        this.listeners.push(callback);
        return callback;
    }
    this.remove = function(callback) {
        for (var i = 0; i < this.listeners.length; i++) {
            if (this.listeners[i] == callback) {
                this.listeners.splice(i, 1);
                return true;
            }
        }
        return false;
    }
    this.dispatch = function() {
        for (var i = 0; i < this.listeners.length; i++) {
            var callback = this.listeners[i];
            try {
                callback();
            }
            catch (err) {
            }
        }
    }
    this.dispatch1 = function(arg1) {
        for (var i = 0; i < this.listeners.length; i++) {
            var callback = this.listeners[i];
            try {
                callback(arg1);
            }
            catch (err) {
            }
        }
    }
    this.dispatch2 = function(arg1, arg2) {
        for (var i = 0; i < this.listeners.length; i++) {
            var callback = this.listeners[i];
            try {
                callback(arg1, arg2);
            }
            catch (err) {
            }
        }
    }
    this.dispatch3 = function(arg1, arg2, arg3) {
        for (var i = 0; i < this.listeners.length; i++) {
            var callback = this.listeners[i];
            try {
                callback(arg1, arg2,arg3);
            }
            catch (err) {
            }
        }
    }
}


var eventsAccountChanged = new WizEventDispatcher();
var eventsClose = new WizEventDispatcher();
var eventsTabCreate = new WizEventDispatcher();
var eventsTabClose = new WizEventDispatcher();
var eventsHtmlDocumentComplete = new WizEventDispatcher();
var eventsDocumentBeforeEdit = new WizEventDispatcher();
var eventsDocumentBeforeChange = new WizEventDispatcher();

function WizOnAccountChnaged() {
    eventsAccountChanged.dispatch();
}


function WizOnClose() {
    eventsClose.dispatch();
}

function WizOnTabCreate(doc) {
    eventsTabCreate.dispatch1(doc);
}

function WizOnTabClose(objHtmlDocument, objWizDocument) {
    eventsTabClose.dispatch2(objHtmlDocument, objWizDocument);
}

function WizOnHtmlDocumentComplete(doc) {
    eventsHtmlDocumentComplete.dispatch1(doc);
}

function WizOnDocumentBeforeEdit(objHtmlDocument, objWizDocument) {
    return eventsDocumentBeforeEdit.dispatch2(objHtmlDocument, objWizDocument);
}
function WizOnDocumentBeforeChange(objHtmlDocument, objWizDocumentOld, objWizDocumentNew) {
    return eventsDocumentBeforeChange.dispatch3(objHtmlDocument, objWizDocumentOld, objWizDocumentNew);
}
```

首先，我们定义了一个class：WizEventDispatcher，用来实现消息的派发。

然后生成了几个事件的实例：

```
var eventsAccountChanged = new WizEventDispatcher();
var eventsClose = new WizEventDispatcher();
var eventsTabCreate = new WizEventDispatcher();
var eventsTabClose = new WizEventDispatcher();
var eventsHtmlDocumentComplete = new WizEventDispatcher();
var eventsDocumentBeforeEdit = new WizEventDispatcher();
var eventsDocumentBeforeChange = new WizEventDispatcher();
```

+ eventsAccountChanged：当切换帐号的的时候触发该事件。该消息在Wiz 3.0里面已经失效。
+ eventsClose：当为知笔记即将关闭的时候触发
+ eventsTabCreate：新开一个tab的时候触发
+ eventsTabClose：关闭一个tab的时候触发
+ eventsHtmlDocumentComplete：当有一个tab里面的html文档加载完成的时候触发
+ eventsDocumentBeforeEdit：当有一个笔记即将被编辑的时候触发
+ eventsDocumentBeforeChange：当某一个tab里面的笔记，即将被另外一个笔记代替的时候触发。

下面这段代码，则真正将为知笔记的事件和插件脚本联系起来：

```
function WizOnAccountChnaged() {
    eventsAccountChanged.dispatch();
}
function WizOnClose() {
    eventsClose.dispatch();
}
function WizOnTabCreate(doc) {
    eventsTabCreate.dispatch1(doc);
}
function WizOnTabClose(objHtmlDocument, objWizDocument) {
    eventsTabClose.dispatch2(objHtmlDocument, objWizDocument);
}
function WizOnHtmlDocumentComplete(doc) {
    eventsHtmlDocumentComplete.dispatch1(doc);
}
function WizOnDocumentBeforeEdit(objHtmlDocument, objWizDocument) {
    return eventsDocumentBeforeEdit.dispatch2(objHtmlDocument, objWizDocument);
}
function WizOnDocumentBeforeChange(objHtmlDocument, objWizDocumentOld, objWizDocumentNew) {
    return eventsDocumentBeforeChange.dispatch3(objHtmlDocument, objWizDocumentOld, objWizDocumentNew);
}
```

当为知笔记某些事件发生的时候，将会调用上面代码中的相应的函数，并传递相应的参数。然后这些函数，将会将这些事件派发给每一个消息响应者。

最后面的部分，简单描述了如何在插件脚本中响应这些事件：


```
/*
function myonclose() {
    WizAlert("onclose");
}
eventsClose.add(myonclose);

function myontabcreate(doc) {
    WizAlert("ontabcreate");
}
eventsTabCreate.add(myontabcreate);

function myontabclose(doc) {
    WizAlert("ontabclose");
}
eventsTabClose.add(myontabclose);

function myonhtmldocumentcomplete(doc) {
    WizAlert("onhtmldocumentcomplete");
}
eventsHtmlDocumentComplete.add(myonhtmldocumentcomplete);
*/
```

例如，要相应为知笔记关闭事件，只需要在自己的插件代码中，添加以下的脚本：

```
function myonclose() {
    WizAlert("onclose");
}
eventsClose.add(myonclose);
```

在用户关闭为知笔记的时候，将会出现一个消息框，显示出 onclose。这时候，插件可以做一些收尾工作，例如保存用户的工作状态。

在某些事件发生的时候，为知笔记会传入相应的参数。

##### 新开一个tab的时候（WizOnTabCreate(doc)）：

doc代表的时tab里面带的html文件的document对象（DOM里面的document对象）。

##### 关闭一个tab的时候（WizOnTabClose(objHtmlDocument, objWizDocument) ）：

objHtmlDocument是tab对应的html文件的document对象；

objWizDocument则对应当前tab打开的为知笔记的笔记对象(IWizDocument，该参数可能为null)。

##### 当为知笔记里面，某一个html文件完成的时候（WizOnHtmlDocumentComplete(doc)）)）：

doc代表的时tab里面带的html文件的document对象（DOM里面的document对象）。

##### 即将编辑某一个笔记的时候（WizOnDocumentBeforeEdit(objHtmlDocument, objWizDocument)））：

objHtmlDocument是tab对应的html文件的document对象；

objWizDocument则对应当前tab打开的为知笔记的笔记对象(IWizDocument，该参数不会为null)。

##### 当某一个tab里面的笔记即将被另外一个笔记代替的时候（WizOnDocumentBeforeChange(objHtmlDocument, objWizDocumentOld, objWizDocumentNew)）：

objHtmlDocument是tab对应的html文件的document对象，

objWizDocumentOld：tab里面原来被打开的笔记对象(IWizDocument)

objWizDocumentNew：tab里面即将被打开的笔记对象(IWizDocument)

例如，当需要相应html打开完整的消息的时候，可以使用下面的代码：

```
function myonhtmldocumentcomplete(doc) {
    WizAlert(doc.title);
}
eventsHtmlDocumentComplete.add(myonhtmldocumentcomplete);
```

当某一个html被打开完整的时候，就会有消息框显示出html文件的标题(通过DOM里面的document.title获得)
