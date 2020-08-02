## 页眉页脚有关操作

您可以通过这个方法快速生成页眉页脚

### 认识

我们可以WordGo对象的 `addHead` 或 `addFoot` 方法把我们的PaperOut添加进去，就可以为文档添加页眉和页脚

### 构造方法

#### 空参构造

`PaperOut()`

不指定任何信息，如果您页眉页脚样式复杂时可以使用。

#### 指定文字构造

`PaperOut(String paperText)`

paperText  页眉或页脚的文字

使用默认样式，适用于只需要在页眉或页脚添加文字的时候使用

#### 指定文字和页眉或页脚全局样式

`PaperOut(String paperText, String paperCss)`

paperText  页眉或页脚的文字

paperCss  页眉或页脚的样式，仅支持border、text-align、font-size属性，使用和属性图片中完全一样。

> 注意：border设置后也不是四周都展示，页眉的话在页眉下面展示一条线，如果是页脚在页脚上方展示一条线



### 方法

添加方法和在WordGo中一样，这里比较简单的介绍一下

#### 添加文字

`addText(String paperText)`

添加文字，样式为默认全局样式



`addCss(String paperCss)`

给页眉或页脚添加一段全局样式，仅支持border、text-align、font-size属性，使用和属性图片中完全一样。



`addStyleText(String styleText, String styleCss)`

添加一段具有独特样式的文本；

styleText  文本内容

styleCss  文本样式，可接受样式和正常添加文字一样

详细支持样式请查看添加文字操作



#### 添加图片

`addImg(String imgUri, String imgStyleCss)`

添加一张图片，仅支持默认流定位，请勿设置为绝对定位或锚点定位(相对定位)；

imgUri  图片绝对路径，支持 jpg、png、gif

imgStyleCss  图片的样式

详细支持样式请查看添加图片操作



#### 添加页码

`addPage()`

添加页码，一个阿拉伯数字表示页码，无特殊样式，不会自动换行。

`addPage(int template, String numStyleCss)`

添加页码并指定模板和样式，numStyleCss样式就和文字样式一样，template支持 1- 3，样式如下：

```
template 取 1：页码
template 取 2：页码 / 总页码
template 取 3：~页码~
```



### 添加到word中

WordGo中有以下两个方法可以将页眉或页脚添加到文档中

`addHead(PaperOut header)`

`addFoot(PaperOut footer)`

分别是添加到页眉和添加到页脚，示例：

示例：

```java
WordCore wordCore = new WordCore("本月经营报告", "王编辑");
//创建时添加文档属性信息
WordGo wordGo = new WordGo(wordCore);
wordGo.add("HelloWorld", "color:#FF0000");
//添加页眉
PaperOut head = new PaperOut("我是页眉");
wordGo.addHead(head);
//添加页脚
PaperOut foot = new PaperOut("我是页脚");
wordGo.addFoot(foot);
//生成文件
wordGo.create("C:\\Users\\Administrator\\Desktop\\1.docx");
```



### 常见问题

> Q：是否支持不同页面不同页眉或页脚
>
> A：抱歉，暂不支持

> Q：页眉插入图片不显示为什么
>
> A：请选择小一些的图片并检查图片路径信息是否正确