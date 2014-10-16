## 日常运维管理

### 软件环境

1. Java环境
    JDK 1.7最新稳定版
1. 数据库
    Mysql 5.5 需要4-5个database，可以放在一个数据库实例中
1. 服务器
    * Apache Tomcat 7.0.47
    * Tengine 2.0.3
    * Nodejs 0.10 稳定版
1. 缓存数据存储
    Redis 2.6 稳定版
1. 搜索服务
    ElasticSearch 0.90.5

### 存储目录结构介绍

* 文件数据：
    1. 笔记数据 /wiz/storage/data_root
    1. 摘要数据 /wiz/storage/abstract_data_root
    1. 索引数据 /wiz/storage/search/es-index
    1. 数据库数据目录位置 /wiz/storage/mysql_databases

* 数据库：
    1. wizasent 用户服务
    1. wizksent 笔记服务
    1. wizkv 存储服务
    1. wiz_message 消息服务

### 服务介绍

#### Java应用

##### as 账号服务

* 服务位置 /wiz/app/wizas
* 日志位置 /wiz/server/apache-tomcat-7.0.37/logs/wizas.log
* 配置文件
    + 应用参数配置 /wiz/app/wizas/WEB-INF/classes/etc/as.properties
    + 数据库参数配置 /wiz/app/wizas/WEB-INF/classes/etc/db.properties
    + 应用配置 /wiz/app/wizas/WEB-INF/classes/etc/config.xml
* 查看是否启动
    ```shell
    curl localhost:8080/wizas/xmlrpc
    # 正常输出：
    XML-RPC server accepts POST requests only.
    ```

##### ks 知识管理服务

* 服务位置 /wiz/app/wizks
* 日志位置 /wiz/server/apache-tomcat-7.0.37/logs/wizks.log
* 配置文件

    + 应用参数配置 /wiz/app/wizks/WEB-INF/classes/etc/ks.properties
    + 数据库参数配置 /wiz/app/wizks/WEB-INF/classes/etc/db.properties
    + 应用配置 /wiz/app/wizks/WEB-INF/classes/etc/config.xml
* 查看是否启动
    ```shell
    curl localhost:8080/wizks/xmlrpc
    # 正常输出
    XML-RPC server accepts POST requests only.
```

##### search 搜索服务

* 服务位置 /wiz/app/wizsearch
* 日志位置 /wiz/logs/ws.log
* 配置文件
    + 应用参数配置 /wiz/app/wizsearch/WEB-INF/classes/application.yml
    + 应用参数配置 /wiz/app/wizsearch/WEB-INF/classes/etc/ws.properties
    + 应用配置 /wiz/app/wizsearch/WEB-INF/classes/etc/app-config.xml
* 查看是否启动
    ```shell
    curl localhost/wizsearch/search
    # 正常输出
    {"return_code":403,"return_message":"非法操作"}
    ```

##### message 消息服务

* 服务位置 /wiz/app/wizmessage
* 日志位置 /wiz/logs/wizmessage.log
* 配置文件
    * 应用参数配置 /wiz/app/wizmessage/WEB-INF/classes/application.yml
    * 应用参数配置 /wiz/app/wizmessage/WEB-INF/classes/etc/config.properties
    * 应用配置 /wiz/app/wizmessage/WEB-INF/classes/etc/app-config.xml
* 查看是否启动
    ```shell
    curl localhost/wizsearch/search
    # 正常输出
    {"result":-1,"return_code":200,"return_message":"success"}
    ```

#### NodeJs应用

##### websdk web端使用的api

* 服务位置 /wiz/app/websdk
* 配置文件
    + /wiz/app/websdk/conf/configCore.js
    + /wiz/app/websdk/conf/config.js

##### manage_console 管理后台

* 服务位置 /wiz/app/manage_console
* 配置文件 /wiz/app/manage_console/config.js
* 日志位置

    日志位于/home/wiznote/.forever
    查看是否启动 使用命令
    ```shell
    forever logs
    ```
    查看日志文件的位置，通过日志文件进行查看

#### API服务

##### api windows客户端使用的api
* 服务位置 /wiz/app/go
* 配置文件 /wiz/app/api/go.lua
* 日志文件位置 /usr/local/nginx/logs/

#### WEB页面

##### note.wiz.cn 前端web页面

静态页面，html、js、css等

* 服务位置 /wiz/app/note.wiz.cn
* 配置文件


#### 配置文件管理

通过puppet管理配置文件 配置文件位置 /wiz/init.pp

可配置参数：

    $companyName：客户端中显示的oem名称
    $openIp: 知识管理服务器的访问ip或域名
    $ip: 127.0.0.1或内网ip（服务分布在多台机器上）
    $dbHost: 数据库地址
    $dbPort: 数据库端口
    $dbUser: 数据库用户
    $dbPwd: 数据库用户的密码

更新所有应用的配置文件的命令

```shell
sudo puppet apply --modulepath=/wiz/EnterpriseDeploy/puppet/modules/ /wiz/init.pp
```

#### Nginx web服务器

反向代理服务器，web服务器

* 查看是否启动
    ```shell
    ps aux|grep nginx
    ```
* 配置文件 /etc/nginx/conf.d/wiz.conf
* 日志文件位置 /var/log/nginx/


#### Mysql 数据库

* 查看数据库是否启动
    ```shell
    ps aux| grep mysql
    ```
* 配置文件 /etc/my.cnf


#### Redis key-value数据存储

存放应用内使用的非持久化数据或缓存数据

* 查看是否启动
    ```shell
    ps aux|grep redis-server
    ```
* 配置文件 /etc/redis-server.conf

#### ElasticSearch 搜索服务器

搜索服务器，供java搜索应用调用，提供搜索功能

* 查看是否启动
    ```shell
    ps aux|grep elasticsearch
    ```
* 配置文件 /wiz/server/elasticsearch-0.90.5/config/elasticsearch.yml
* 日志文件 /wiz/server/elasticsearch-0.90.5/logs/elasticsearch.log

### 运维建议

为保障企业私有部署的正常运行，数据的安全，建议建立常规的服务巡检制度和数据备份制度。

定时对所有服务进行巡检，查看磁盘使用情况、各个服务运行情况，以保障企业私有部署服务的正常稳定运行。

定时对备份数据进行检查，查看数据是否正常备份，以应对由于硬件损坏、自然灾害而出现数据丢失的情况。

如果出现服务器错误导致无法登录，或者某些功能不正常，需要对所有服务进行检查，查看服务是否正常运行。如果服务因错误而中止，需要重新将其启动。

### 数据备份

备份需要备份两部分数据，数据库和文件

#### 数据库备份

数据库备份可以使用sqldump命令备份整个数据库，在服务器的终端中执行该命令会将数据库保存到当前目录的mysql_databases.bak文件中

```
mysqldump -uroot -proot --databases wizasent wizksent wiz_message wizkv > mysql_databases.bak
```

#### 文件备份

文件备份需要备份整个/wiz/storage目录
