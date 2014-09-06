## 为知笔记帮助手册

### 打包和nginx设置

由于改动了配置文件结构，无法使用官方的gitbook进行打包，请使用我们修改过的支持中文搜索的gitbook版本

npm install -g codepiano/gitbook

打包后的文件在 repo路径/manual/ 下，每次生成网站都会先清空manual目录，有一个简陋的index.html，方便测试时预览

```bash
sh pkg.sh
```

nginx设置,添加

```
location / {
    root   repo路径/manual;
    index  index.html;
}


location ^~ /manual {
    root   repo路径/;
    index  index.html;
}

```
### favicon.ico

运行时会尝试分析出gitbook的路径，然后使用目录中的替换gitbook默认的favicon.ico

### markdown文件规范

1. 按照markdown语法规范
1. 尽量不要使用html标签，目前只有锚点无法使用纯markdown实现
1. 使用-分割文件名中的单词
1. 不要出现多余的空行、空白字符
1. 图片命名规范化，旧的命名没必要浪费精力修改了，以后更新的时候尽量使用有意义的名称

### 所用插件

在gitbook所在目录下安装

1. 第三方评论插件，集成多说评论
    ```shell
    npm isntall gitbook-plugin-duoshuo
    ```
 
1. 谷歌分析插件
    ```shell
    npm install gitbook-plugin-ga
    ```

