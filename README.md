## 为知笔记帮助手册

### 打包和nginx设置

每次都会先清空preview目录，有一个简陋的index.html，方便测试时预览

```
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
