## 设置纸张大小和边距

我们有一个专门的类 WordPaper 用来设置纸张。

### 认识

我们可以创建一个wordPaper对象，在创建wordGo的时候通过构造方法传入。（详见 WordGO 构造和生成文件 ）

```java
WordPaper wordPaper = new WordPaper("A3");
WordGo wordGo = new WordGo(wordPaper);
```

### 构造方法

WordPaper 构造方法提供了丰富的操作，不使用set操作，就可以创建出来一个您想要的 wordPaper 对象。来选一个适合您的吧：

#### 仅指定纸张类型

`WordPaper(String paper)`

paper为可接受的纸张类型，详细见文末对应表格。默认为竖向，默认页边距*。

❤ 如果您只需要指定纸张类型可以直接在WordGo创建时传递paper类型，效果一样哦~~

> `*` 默认页边距指：上下2.54cm、左右3.18cm。该边距是Office软件的默认边距，不建议修改~



#### 指定纸张类型和方向

`WordPaper(String paper, boolean landscape)`

paper  为可接受的纸张类型，详细见文末对应表格。

landscape  为方向，true为横向，false为竖向。

页边距为默认页边距。



#### 全部尺寸都自定义

`WordPaper(int paperWidth, int paperHeight, int paddingTop, int paddingRight, int paddingBottom, int paddingLeft, int paperHeader, int paperFooter, boolean landscape)`

| 参数          | 含义                                  |
| ------------- | ------------------------------------- |
| paperWidth    | 纸张宽度，单位mm                      |
| paperHeight   | 纸张高度，单位mm                      |
| paddingTop    | 页面上边距，单位mm                    |
| paddingRight  | 页面右边距，单位mm                    |
| paddingBottom | 纸张下边距，单位mm                    |
| paddingLeft   | 纸张左边距，单位mm                    |
| paperHeader   | 页眉顶端距离，单位mm                  |
| paperFooter   | 页脚底端距离，单位mm                  |
| landscape     | 纸张方向方向，true为横向，false为竖向 |

页眉顶端距离 和 页脚底端距离 不太常用，一般分别设置成150和175，Office可以在页眉页脚中找到这一项。

![页眉页脚距离](C:\Users\Administrator\Desktop\页眉页脚距离.png)



### 设置分栏

`colsAverage(int num, double space, boolean sep)`

num 需要的分栏数量

space 每两个分栏之间间隔的字符数，office推荐值2.02

sep 是否需要分栏线，true需要，false不需要



### 支持纸张关系表

纸张单词之间可以什么都不写或者使用 “-” 或 “_” 连接，不区分大小写。如果您传递了一个没有预设的纸张类型将会抛出一个RuntimeException

| 接收值    | Office名称 | 执行标准 | 大小/mm       |
| --------- | ---------- | -------- | ------------- |
| A3        | A3         | ISO 216  | 297 * 420     |
| A4        | A4         | ISO 216  | 210 * 297     |
| A5        | A5         | ISO 216  | 148 * 210     |
| B4JIS     | B4(JIS)    | JIS      | 257 * 364     |
| B5JIS     | B5(JIS)    | JIS      | 182 * 257     |
| tabloid   | tabloid    | ANSI     | 279.4 * 431.8 |
| statement | statement  | 不详     | 139.7 * 215.9 |
| executive | executive  | ANSI     | 184.1 * 266.7 |
| letter    | 信纸       | ANSI     | 215.9 * 279.4 |
| legal     | 法律专用纸 | ANSI     | 215.9 * 355.6 |
| 16cut     | 16开       | 国内常用 | 184 * 260     |
| 32cut     | 32开       | 国内常用 | 130 * 184     |
| big32cut  | 大32开     | 国内常用 | 140 * 203     |

示例：`WordPaper wordPaper = new WordPaper("A3")`