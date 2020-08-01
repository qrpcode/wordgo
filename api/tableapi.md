## 设置和添加表格有关操作

### 写在最前

表格计算行列时，全部和我们平时使用一样，**从1开始**



### 创建表格

`new WordTable(int row, int column)`

`new WordTable(int row, int column, String css)`

创建一个 row 行，column 列的表格，支持在创建的时候设置css样式

> 这里css支持直接指定模板，省时省力！详见下方 template 属性



### 合并单元格

`merge(int rowLeftTop, int columnLeftTop, int rowRightBottom, int columnRightBottom)`

rowLeftTop、columnLeftTop 分别为合并区域的左上单元格行和列；rowRightBottom、columnRightBottom分别为合并区域的右下单元格行和列。

合并成功返回true，失败返回false。

合并后会将所有被合并单元格内容拼接。

代码示例：

```java
WordTable wordTable = new WordTable(5, 6, "column-width:1=50%; template: normal2; width:50%");
wordTable.merge(2, 4, 4 ,5);
```

成功和失败示例：

![单元格合并](C:\Users\Administrator\Desktop\单元格合并.png)

### 填充数据

#### 添加文本

`add(int row, int column, String ... str)`

从 row 行column列开始向左填充数据。

如果遇到合并单元格且是左上角单元格才会填充（详见示例）

示例1（左）：`wordTable.add(2, 1, "aa", "bb", "cc")`

示例2（右）：`wordTable.add(3, 1, "aa", "bb", "cc")`

![合并左上用](C:\Users\Administrator\Desktop\合并左上用.png)

#### 添加图片

`addImg(int row, int column, String uri, String css)`

在 row 行column列中插入一个绝对地址为uri的图片，样式表为css（css样式可用范围和正常插入图片一样）。成功返回true，失败返回false。



#### 添加独特样式的文本

`addStyleText(int row, int column, String text, String css)`

在 row 行column列中插入 text 文字，样式表为css（css样式可用范围和正常插入文字一样）。成功返回true，失败返回false。



#### 设置部分单元格样式

设定某一个：`addStyle(int row, int column, String css)`

设定一些：`addStyle(int rowLeftTop, int columnLeftTop, int rowRightBottom, int columnRightBottom, String css)`

设定一些的前四个参数和合并单元格同理，是左上和右下单元格的行和列值，css为样式表文本。



### 支持样式

图片和文字样式和正常使用一样，这里仅介绍表格样式支持的样式

#### template 表格模板

表格模板和MSoffice一样，下面是名称和对应样式对照关系，此模板支持WPS等打开浏览。

示例：`template:normal3`

![表格列表](C:\Users\Administrator\Desktop\表格列表.png)

详细样式是什么样子可以在Microsoft Word查看：

![表格样式](C:\Users\Administrator\Desktop\表格样式.png)



#### width 表格宽度

支持百分比，也支持磅值。百分比必须写%，写pt、px或单独一个数字都会被解析为磅值

示例：`width:50%`  `width:300pt`



#### text-align 对齐方式

和文字一样，left、center、right三个取值，推荐center，默认left

示例：`text-align:center`  `text-align:left`



#### column-width 列宽

需要指定哪一列多宽，支持百分比或磅值。百分比需带%，纯数字或者px、pt结尾都会被解析为磅值

`column-width:1=50%,3=20%`  

如上，用“ , ” 分割不同列，写法是列号 = 宽度

其他未指定的列将会自动平分宽度。

> 【重要】即使指定了宽度，如果在格子中插入了图片，桌面版office软件一般都还会自动调整，插入图片后，可能设置的值和实际仍有偏差。石墨文档等在线工具导入一般不会进行调整。



#### row-height 行高（不推荐使用）

> 【重要】使用此属性可能导致文字因为高度太小显示不全！默认为自动调整，建议保持默认。

示例：`row-height:1=50pt,3=20%`

使用方法同 column-width



