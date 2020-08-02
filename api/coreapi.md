## 设置文件作者和摘要

如果您没有设置默认为空，我们也不是很建议您进行设置，很多软件都 **根本看不到这些信息**。

### 认识

我们可以创建一个wordPaper对象，在创建wordGo的时候通过构造方法传入。（详见 WordGO 构造和生成文件 ）

```java
WordCore wordCore = new WordCore("本月经营报告", "王编辑");
WordGo wordGo = new WordGo(wordCore);
```

### 构造方法

wordCore 构造方法提供了丰富的操作，不使用set操作，就可以创建出来一个您想要的 wordPaper 对象。

#### 指定标题

`WordCore(String title)`

title  文章标题信息

#### 指定标题和作者

`WordCore(String title, String creator)`

title   文章标题信息

creator  文章作者信息

#### 指定全部

`WordCore(String title, String subject, String creator, String keywords, String description, String lastModifiedBy, int revision, Date created, Date modified)`

| 参数           | 含义         |
| -------------- | ------------ |
| title          | 文档标题     |
| subject        | 标记         |
| creator        | 作者         |
| keywords       | 标记         |
| description    | 备注         |
| lastModifiedBy | 上次修改者   |
| revision       | 经理         |
| created        | 创建时间     |
| modified       | 上次修改时间 |

含义名称和office一样，详细含义请自行打开参考~

Microsoft Office查看：

![msoffice查看属性](https://github.com/qrpcode/wordgo/blob/master/api/textapi.assets/msoffice%E6%9F%A5%E7%9C%8B%E5%B1%9E%E6%80%A7.png?raw=true)

WPS查看：

![Wps查看文件属性](https://github.com/qrpcode/wordgo/blob/master/api/textapi.assets/Wps%E6%9F%A5%E7%9C%8B%E6%96%87%E4%BB%B6%E5%B1%9E%E6%80%A7.png?raw=true)

