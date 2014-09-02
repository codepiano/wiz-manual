## 安装和使用中常见问题

1. 安装过程出错，执行终止

    可能有两个原因。
    + 系统版本不一致，导致缺少了部分软件包。请使用官方镜像进行安装，系统版本要完全一致。
    + 安装的时候选择了最小化安装（minimal mode），安装操作系统的时候选择该模式也会导致缺少软件包的情况。

    由于安装过程为非可重复性操作，如果出错，建议重新安装要求的操作系统，再次进行私有部署的安装。

1. tomcat无法正常启动

    系统hosts配置有误，未将本机的机器名解析到127.0.0.1

1. 安装以后可以打开登录页面，但无法正常登录

    检查as服务是否启动，检查方法见<a href="./administrator.html">运维相关</a>

1. WEB端可以正常登录和新建笔记，Windows客户端无法正常同步

    检查windows客户端的安装目录是否有名为 oem.ini 的文件，检查该文件的内容，里面的ip地址设置是否与私有部署服务器的访问ip一致