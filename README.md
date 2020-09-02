# Word GO

WordGO - 让Java生成word文档更容易.

【choose language：[English](https://github.com/qrpcode/wordgo/blob/master/README_EN.md)】

### 项目背景

传统的Java生成word通常需要先手动创建模板文件，之后导入。如果不希望创建模板，还想少些点代码，选Word GO是个好主意~~

### 安装

#### 手动导入jar包

* IDEA导入：点击File-Project Structure；然后在左侧找到Modules并点击；最后在右侧点击绿色的+号，选择JARs or directories选取要导入的jar包即可。
* Eclipse导入：右击“项目”→选择Properties，在弹出的对话框左侧列表中选择Java Build Path

#### maven中央仓库导入

```xml
<dependency>
 <groupId>com.github.qrpcode</groupId>
 <artifactId>wordgo</artifactId>
 <version>1.0-SNAPSHOT</version>
</dependency>
```

### 环境依赖和兼容性

只要能运行java这个就能用，他不依赖于任何第三方Office应用

兼容性请看表：

| 分类     | 软件名称         | 兼容性说明                                                   |
| -------- | ---------------- | ------------------------------------------------------------ |
| 桌面版   | *Microsoft* Word | 完美兼容                                                     |
|          | WPS              | 完美兼容                                                     |
|          | 永中Office       | 完美兼容                                                     |
| 在线版   | 金山文档         | 该系统本身对部分属性不支持                                   |
|          | 腾讯文档         | 该系统本身对部分属性不支持                                   |
|          | 石墨文档         | 该系统本身对部分属性不支持，偶尔可能提醒文档需要修复，但能正常导入和显示 |
| 其他常用 | 手机QQ           | 该系统本身对部分属性不支持                                   |
|          | 手机微信         | 该系统本身对部分属性不支持                                   |

[注] 经过测试，石墨文档应该不是用xml解析的，似乎是字符串截取的，正在努力适配~



### 使用

来，导入了jar包，我们先来创建一个“Hello World”

```java
WordGo wordGo = new WordGo();
//新建一个word
wordGo.add("Hello World", "font-size: 15; color: #FF0000");
//填充数据可以查看对应功能说明
wordGo.create("C:\\demo.docx");
//最后生成即可，参数是生成目录，必须带文件名且以.docx结尾
```

是的，它和Css写法很类似，很容易上手~~

代码支持JDK1.5 +（含）

#### 对应功能说明（10分钟就能学会）

👉 [WordGO 构造和生成文件](https://github.com/qrpcode/wordgo/blob/master/api/wordgoapi.md)

👉 [设置文档属性信息](https://github.com/qrpcode/wordgo/blob/master/api/coreapi.md)

👉 [设置纸张大小和边距](https://github.com/qrpcode/wordgo/blob/master/api/paperapi.md)

👉 [文字、换页有关操作](https://github.com/qrpcode/wordgo/blob/master/api/textapi.md)

👉 [创建、填充、添加表格有关操作](https://github.com/qrpcode/wordgo/blob/master/api/tableapi.md)

👉 [图片有关操作](https://github.com/qrpcode/wordgo/blob/master/api/imgapi.md)

👉 [页眉页脚有关操作](https://github.com/qrpcode/wordgo/blob/master/api/paperoutapi.md)

### 主要项目负责人

[@qrpcode](https://github.com/qrpcode)

### 参与

没错，我也觉得我代码写的 ~~有点~~ (十分) 乱

来帮帮我吧，Fork 之后 pull request 一下就可以啦~

### 开源协议

[Apache-2.0 License](https://github.com/qrpcode/wordgo/blob/master/LICENSE)

（也就是说他是可以商用的，详细看协议吧~~）



### 💖 如果觉得有用记得点 Star⭐

.



### 当前仍为快照版，还存在较多bug，

### <u>不建议用于生产环境</u>

发现BUG随时发邮件到  i@qiruipeng.com  我会尽快回复和修复的哟~~

> 已发现BUG：
>
> 1. 样式表只写一个且属性为颜色的时候可能会无效
> 2. addLine方法逻辑存在问题，部分时候可能无法正常换行