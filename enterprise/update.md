## 如何升级企业私有部署

1.  停止tomcat服务
执行下面命令，获取tomcat进程号
```shell
ps aux | grep tomcat | awk '{print $2, $11}'
```
结果类似
```shell
1234 /usr/bin/java
12345 grep
```
在输出结果中找到 /usr/bin/java进程号，上例中是1234，然后执行
```shell
kill 1234
```
杀掉tomcat进程
_每次的进程号都不一样，需要先找出进程号，再杀掉进程_

    _2014-8-11发布的版本，jdk版本升级到了7.0，如果需要升级私有部署，停止tomcat以后，需要升级jdk版本，具体方法见下面的<a href="#other">其他更改</a>_

1. 备份旧的app目录
执行下面命令备份旧的app目录
```shell
mv /wiz/app /wiz/app`date "+%Y-%m-%d_%H:%M:%S"`
mv /wiz/init.pp /wiz/init.pp`date "+%Y-%m-%d_%H:%M:%S"`
```

1. 获取最新的私有部署升级包
从百度网盘中获取私有部署升级包wizupdate.tar.gz，地址：http://pan.baidu.com/s/1u8Trw
将目录app和文件init.pp放入 /wiz
用目录EnterpriseDeploy替换旧的 /wiz/EnterpriseDeploy
修改 /wiz/init.pp 中的设置，修改openIp等相关信息
执行命令修改配置文件
```shell
sudo puppet apply --modulepath=/wiz/EnterpriseDeploy/puppet/modules/ /wiz/init.pp
```
正常情况下的输出为
```shell
XML-RPC server accepts POST requests only.
XML-RPC server accepts POST requests only.
```

1. 更新数据库
```shell
sh database_update.sh
```
输入密码后执行更新脚本

1. 启动tomcat
```shell
/wiz/server/apache-tomcat-7.0.37/bin/startup.sh
#检查tomcat是否正常启动
curl localhost:8080/wizas/xmlrpc && curl localhost:8080/wizas/xmlrpc
```

1. 重启nodejs应用
```shell
forever restartall
```

1. 更新完成，检查网页版和各客户端是否正常

<a name="other" />其他更改：
+ 2014-8-11发布的版本，jdk版本升级到了7.0，以前的用户如果需要升级私有部署，停止tomcat以后，需要先升级jdk版本
     可以在oracle官网或者下面地址获取jdk-7u65-linux-x64.rpm ：http://yun.baidu.com/s/1u8Trw
     下载完成后执行
     ```shell
     sudo rpm -ivh jdk-7u65-linux-x64.rpm
     ```
