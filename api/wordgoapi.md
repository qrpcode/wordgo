## WordGO 构造和生成文件

您可以创建 WordGO 对象后，添加表格、图片文字等，最后可以生成一个文档文件~

### 创建文件

#### 空参默认值构造 🔥 

`WordGo()`

满足您的一般需求，默认A4纸竖向，页边距和页眉页脚距离全部采用office相同默认尺寸。

#### 指定纸张类型构造

`WordGo(String paperType)`

paperType：纸张类型，详细请查看WordPaper类介绍中可接受属性表格。

示例：`new WordGo("A3")`

#### 指定纸张信息构造

`WordGo(WordPaper wordPaper)`

wordPaper  纸张类，具体创建请参考WordPaper类介绍。

#### 指定文档属性信息构造

`WordGo(WordCore wordCore)`

wordCore  文档属性对象，详细请参考WordCore介绍。

#### 指定文档属性信息和纸张构造

`WordGo(WordCore wordCore, WordPaper wordPaper)`

wordCore  文档属性对象，详细请参考WordCore介绍；

wordPaper  纸张对象，具体创建请参考WordPaper介绍。



### 添加内容方法

👉 [文字、换页有关操作](https://github.com/qrpcode/wordgo/blob/master/api/textapi.md)

`addLine(String text, String css)`

`add(String text, String css)`

`newPage()`

👉 [创建、填充、添加表格有关操作](https://github.com/qrpcode/wordgo/blob/master/api/tableapi.md)

`add(WordTable wordTable)`

`addTable(WordTable wordTable)`

👉 [图片有关操作](https://github.com/qrpcode/wordgo/blob/master/api/imgapi.md)

`addImg(String uri, String css)`

👉 [页眉页脚有关操作](https://github.com/qrpcode/wordgo/blob/master/api/paperoutapi.md)

`addHead(PaperOut header)`

`addFoot(PaperOut footer)`



### 生成文件

`create(String fileWay)`

fileWay  指定生成的目录和文件名，文件名必须以  .docx  结尾，路径必须为绝对路径

示例：`wordGo.create("C:\\HelloWorld.docx");`



### 常见问题

> Q：为什么一开始就需要定义纸张信息，可以中途设置或修改吗？
>
> A：不可以，当图片表格等距离或大小设置为百分比需要通过纸张大小求出实际大小，不提前指定或中途更改会无法计算。

