## 为知笔记帮助手册

### 打包 

每次都会先清空preview目录，有一个简陋的index.html，方便测试时预览

```
sh pkg.sh
```

### favicon.ico

运行时会尝试分析出gitbook的路径，然后使用目录中的替换gitbook默认的favicon.ico
