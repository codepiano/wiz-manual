Wiz 从1.06 开始，更改了新建模板的方式，允许以填表的形式新建各种文档并进行保存，具体的例子，可以参看系统内置的 日记模板和每日回顾模板。

### 新建文档模板是什么格式？

新建文档模板，就是一个HTML文件。您可以用html制作表单，然后在用户保存的时候，将用户输入的内容重新组织成一个HTML文件，保存到Wiz里面。

#### 一些例子

+ 新建日记的模板：Wiz安装目录下面的templates\new\journal.htm文件
+ 每日回顾模板：Wiz安装目录下面的templates\new\daily_review.htm文件

您可以使用文本编辑器或者html编辑器打开这些文件（不要用word/wps之类的字处理软件打开）。

例如我用Visual Studio 编写了一个html文件，如下图：

![html模板](../img/template-html.jpg)

html源代码：

```
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title></title>
        <style type="text/css">
            #TextArea1 {
                height: 200px;
                width: 508px;
            }
            #Text1 {
                width: 470px;
            }
        </style>
    </head>
    <body>
        <p>Title: <input id="Text1" type="text" /></p>
        <p>Text:</p>
        <p><textarea id="TextArea1" name="S1"></textarea></p>
    </body>
</html>
```

### 新建文档模板如何被Wiz找到？

新建文档的模板，必须放在Wiz安装目录下面的 templates\new这个文件夹内，扩展名必须是.htm

将前面编辑的htm文件保存到templates\new，名称为：my_wiz_template.htm。然后重新启动Wiz，点击工具栏的新建按钮右边的三角，会出现菜单：

![菜单](../img/template-menu.jpg)

可以看到，这里显示的是一个英文的菜单，也就是文件名。

为了显示出中文，或者自动适应多语言环境，进行本地化，我们建立一个my_wiz_template.ini文件（和my_wiz_template.htm文件保存在一个文件夹内），内容如下：

```
[Strings]
my_wiz_template=My Wiz Template
my_wiz_template_2052=我的Wiz模板
my_wiz_template_1028=我的Wiz模板
```

注意：

在[Strings]部分，定义了字符串翻译的方法。其中等号左边是字符串的名称，右边是翻译的结果。

名称后面如果包含_xxxx（xxxx是一个数字），则表示是某一种语言下面的翻译。

其中2052是简体中文呢，1028是繁体中文，具体可以查询LCID的值。

在新建文档菜单中，显示的名称，会通过文件名到ini文件里面查找具体的翻译结果，在这个例子中，我们的文件名是my_wiz_template，因此字符串翻译的时候，去查找my_wiz_template_2052，获得了“我的Wiz模板”这个字符串，用来显示在菜单中。

注意ini文件保存的时候，应该使用unicode或者utf-8(带签名)编码。

![编码](../img/template-encoding.jpg)

重新启动Wiz，可以看到菜单已经是中文的了。

![编码](../img/template-cn.jpg)

点击这个菜单，可以看到一个对话框：

![编码](../img/template-dialog.jpg)

### 如何保存用户输入内容？

用户在这个对话框内输入内容后，需要保存。这个时候我们需要增加一个保存按钮，明且相应用户点击的消息。

编辑网页，增加一个按钮，然后给它添加onclick消息：

```
<input id="Button1" type="button" value="Save" onclick="return Button1_onclick()" /></p>

<script language="javascript" type="text/javascript">
    function Button1_onclick() {

    }
</script>
```

下面我们编写脚本：

```
<script language="javascript" type="text/javascript">
var objApp = window.external; //window.external对象，就是WizExplorerApp这个对象，用来和WizExplorer来交互
var objDatabase = objApp.Database; //当前用户数据库
var objWindow = objApp.Window; //当前窗口
var objCategoryCtrl = objWindow.CategoryCtrl; //文件夹列表控件
var objCurrentFolder = objCategoryCtrl.SelectedFolder; //当前选中的文件夹
if (objCurrentFolder == null) { //如果没有选中
objCurrentFolder = objDatabase.GetFolderByLocation("/My Drafts/", true); //获得我的草稿这个文件夹
}
//
var objDoc = null; //生成的WizDocument对象
//
function saveToWiz() { //保存到Wiz
try {
var title = Text1.value; //获得用户输入的标题
if (title == null || title == "") {
title = "Untitled";
}
//
if (objDoc == null) { //如果是第一次保存
objDoc = objCurrentFolder.CreateDocument2(title, ""); //生成文档
objDoc.Type = "mywiztype"; //设置文档类型
}
//
objDoc.ChangeTitleAndFileName(title); //更改文档标题和文件名
//
var htmltext = "<h1>" + title + "</h1><div>" + TextArea1.innerHTML + "</div>"; //生成的文档的html内容
//
//
objDoc.UpdateDocument3(htmltext, 0); //将html内容保存到文档中
//
objWindow.CategoryCtrl.SelectedFolder = objCurrentFolder; //选中当前文件夹
objWindow.DocumentsCtrl.SetDocuments(objCurrentFolder); //让文档列表显示当前文件夹的文档列表
objWindow.DocumentsCtrl.SelectedDocuments = objDoc; //选中当前文档
objWindow.ViewDocument(objDoc, true); //在WizExplorer里面显示保存后的文档
}
catch (err) {
alert(err.message); //出错了
}
}
//
function Button1_onclick() {
saveToWiz();
}
</script>
```

然后将html文件保存，重新使用这个模板，新建文档，就可以点击保存按钮进行保存了。

效果就是将用户输入的标题和文字保存到当前文件夹下面的新文档中。

### 如何本地化？

从上面的例子中，可以看到显示的新建模板对话框里面的html没有进行本地化。

新建模板文档的本地化，和html对话框插件的本地化方法一样，将会在稍后介绍。

### 例子下载：
请下载附件，然后把zip解压缩，将所有的文件复制到Wiz安装文件夹的templates\new这个文件夹内，重新启动Wiz测试。

