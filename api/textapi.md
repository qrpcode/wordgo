## 添加文字

### 添加方法

#### add(String text, String css)

添加一段带有样式的文本

text：文本内容，\n 或 \r会被视为换行（直接写，无需添加\）

css：样式表内容，支持的样式后面介绍

```java
WordGo wordGo = new WordGo();
//创建一个wordGo对象
wordGo.add("Hello", "font-size: 20; background-color: blue;");
//添加文字“Hello”，字号用20号，背景颜色选择blue
wordGo.create("C:\\demo.docx");
//将docx文件输出到C盘根目录下，命名为demo.docx
```



#### addLine(String text, String css)

添加文本并且在最后添加一个换行表示。实际上底层就是调用额add方法，可以省略最后的\n而已。

示例：

```java
WordGo wordGo = new WordGo();
//创建一个wordGo对象
wordGo.addLine("Hello", "font-size: 15; color: #FF0000;");
//添加一行文字“Hello”，字号用15号，颜色选择红色
wordGo.create("C:\\demo.docx");
//将docx文件输出到C盘根目录下，命名为demo.docx
```



### 换页方法

#### newPage()

切换到新的一页，不管这一页写没写完。

```java
WordGo wordGo = new WordGo();
wordGo.add("page1", "");
wordGo.newPage();
//直接跳转到下一页
wordGo.add("page2", "");
wordGo.create("C:\\demo.docx");
```



### 支持样式

#### font-size 文字大小

设定文字大小，支持数字，也支持我们常用的表示字号文字：

> "初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号","小五", "六号", "小六", "七号", "八号"

数字不需要带单位，如果希望知道数字对应多大可以打开word试一下（如图）。无法识别的font-size将会设置成默认的五号字

示例：`font-size:12`  `font-size:五号`

![字号](D:\hello\wordgo\api\textapi.assets\字号.png)



#### font-family 字体

直接写字体名称即可，如果打开的用户电脑有这个字体将会展示，没有的话将会展示默认字体。WPS可能会帮助用户自动下载或提示替换。

示例：`font-family:宋体` `font-family:思源黑体`

![字体](D:\hello\wordgo\api\textapi.assets\字体.png)



#### font-weight 文字加粗

支持normal和bold，不写此属性默认为normal

示例：`font-width:bold`  `font-width:normal`



#### font-style 字体样式

支持normal、italic、oblique。normal是默认不倾斜样式，italic 和 oblique都会让文字倾斜，底层实现没有差别，只是为了照顾不同css书写习惯用户。

示例：`font-style:italic`  `font-style:oblique`  



#### color 文字颜色

文字颜色属性，仅支持16位表示法，不支持带透明度。# 带不带都行。

示例：`color:#FF0000`  `color:#008000`  `color:00FF00`



#### background-color 文字背景色

注意，word中文字背景色仅支持固定几种（如图），受支持的颜色有（顺序如图）：

> "yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue","darkCyan", "darkGreen", "darkMagenta", "darkRed", "darkYellow", "darkGray", "lightGray", "black"

示例：`background-color:yellow`   `background-color:red`

![背景色](D:\hello\wordgo\api\textapi.assets\背景色.png)



#### text-decoration 文字下划线

文字下划线支持多种样式，不写则为没有下划线。划线样式不区分大小写。支持的样式有：

> "single", "double", "thick", "dotted", "dash", "dotDash", "dotDotDash","wave"

示例：`text-decoration:single`  `text-decoration:double`

对应关系：

![下划线](D:\hello\wordgo\api\textapi.assets\下划线.png)



#### text-decoration-color 文字下划线颜色

如果设定了下划线，默认和文字一个颜色。如果设置了独特颜色当然也被支持。仅支持16位表示法，不支持带透明度。# 带不带都行。

示例：`text-decoration-color:#FF0000`  `text-decoration-color:#008000`  `text-decoration-color:00FF00`

![下划线颜色](D:\hello\wordgo\api\textapi.assets\下划线颜色.png)



#### font-emphasis 文字着重号

他用来表示文字下方有小圆点着重强调。支持normal和point属性，当属性值为point表示有着重号，不写默认为normal，即没有着重号。

示例：`font-emphasis:point`  `font-emphasis:normal`

![着重号](D:\hello\wordgo\api\textapi.assets\着重号.png)

#### font-scale 文字水平缩放

文字的水平缩放比率，支持0~100（不支持0，不写默认为100）的整数。写不写%都可以解析。

示例：`font-scale:90`  `font-scale:50%`

![缩放](D:\hello\wordgo\api\textapi.assets\缩放.png)



#### line-height 行高

行高是整行文本的整体属性，建议只设置一个，设置多个我们也无法保证解析的是哪一个。

支持百分比形式和固定磅值形式。百分比必须带%，固定磅值支持以px或pt为单位（都会解析为磅，防止单位写错解析不了）。只有数字不写单位视为固定值。固定值支持 1 - 500（含）磅。

示例：`line-height:150%`  `line-height:18pt`

![行距](D:\hello\wordgo\api\textapi.assets\行距.png)



#### text-align  对齐方式

支持的参数有：left（左对齐）、right（右对齐）、center（居中对齐），默认为左对齐

示例：`text-align:center`

![对齐](D:\hello\wordgo\api\textapi.assets\对齐.png)



### 常见问题

> Q：错误的css是否会被解析报错？
>
> A：错误的css不会被解析，会被直接忽略，和html一样，很难写出错误

> Q：css分割符号和html是否一样？
>
> A：一样的，示例：text-align: center;

> Q：一个属性重复定义会以最后一个为准吗？
>
> A：不一定，只能保证解析出来的效果是您定义中的一个，是哪个无法确定

