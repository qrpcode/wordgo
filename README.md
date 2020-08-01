### 代码和jar包、api正在整理，将在最近发布

### 感谢关注！















# Word GO

WordGO - 让Java生成word文档更容易

### 项目背景

传统的Java生成word通常需要先手动创建模板文件，之后导入。如果不希望创建模板，还想少些点代码，选Word GO是个好主意~~

### 安装

目前仅支持手动导入jar包，准确的来说我正在申请maven收录

### 环境依赖和兼容性

只要能运行java这个就能用，他不依赖于任何第三方Office应用

兼容性请看表：

![我的兼容性](https://gitee.com/qiruipeng/qiruipeng/raw/master/img/jianrong.png)

### 使用

来，我们先来创建一个“Hello World”

```java
WordGo wordGo = new WordGo();
wordGo.add("Hello World", "font-size: 15; color: #FF0000; background-color: blue;");
wordGo.addFoot(foot);
```

是的，它和Css写法很类似，很容易上手~~

代码支持JDK1.5 +（含）

#### 对应功能说明（10分钟就能学会）

👉 设置文件作者和摘要

👉 设置纸张大小和边距

👉 添加文字有关操作

👉 设置和添加表格有关操作

👉 添加图片有关操作

👉 页眉页脚有关操作

### 主要项目负责人

[@qrpcode](https://github.com/qrpcode)

### 参与

没错，我也觉得我代码写的 ~~有点~~ (十分) 乱

来帮帮我吧，Fork之后pull request一下就可以啦~

### 开源协议

[Apache-2.0 License](https://github.com/qrpcode/wordgo/blob/master/LICENSE)

（也就是说他是可以商用的，详细看协议吧~~）



### 如果觉得有用记得点 Star⭐