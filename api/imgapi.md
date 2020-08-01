## 添加图片

### 添加图片方法

#### addImg(String uri, String css)

uri：本地图片绝对路径。仅支持jpg、png、gif图片，在word中，bmp等图片都会被转码成为png储存，所以如有需要保存其他格式请自行转码~~

css：图片的样式

示例：（运行demo前记得修改图片绝对路径❤）

```java
WordGo wordGo = new WordGo();
wordGo.addImg("C:\\Users\\Administrator\\Desktop\\image.png", "border: 3px solid #FF0000; box-shadow: out-bottom-right; new-line:true; text-align:center");
wordGo.addImg("C:\\Users\\Administrator\\Desktop\\image2.gif", "reflection: near-max; new-line:true; position: fixed; left: 50%; top:70%;");
wordGo.create("C:\\demo.docx");
```



### 支持样式

#### width 图片宽度

设置图片宽度，单位为px

#### height 图片高度

设置图片高度，单位为px



#### new-line 独立一行显示

属性被设置为true，会独立成行。

如果您设置了图片文档定位或者相对于锚点定位会自动设置为true（详见下一项）

示例：`new-line:true`  

![独立成行](https://github.com/qrpcode/wordgo/blob/master/api/textapi.assets/%E7%8B%AC%E7%AB%8B%E6%88%90%E8%A1%8C.png?raw=true)



#### position 定位类型

absolute：锚点定位，相对于插入点所在行靠左边界定位

fixed：绝对定位，相对于文档可使用区域（即去除页边距后）左上角定位

static：默认值，没有定位，会出现在正常的流中

注意：fixed绝对定位只能在文档第一页使用

示例：`position:absolute`  `position:fixed`

![图片定位](https://github.com/qrpcode/wordgo/blob/master/api/textapi.assets/%E5%9B%BE%E7%89%87%E5%AE%9A%E4%BD%8D.png?raw=true)



#### box-shadow 图片阴影

只支持office默认的一些样式，对应属性关系和样式请参考下图

![阴影关系](https://github.com/qrpcode/wordgo/blob/master/api/textapi.assets/%E9%98%B4%E5%BD%B1%E5%85%B3%E7%B3%BB.png?raw=true)

中间间隔符号使用 "-" 或 "_" 或 不写都可以正常解析。不区分大小写。

示例：`box-shadow:out-right`  `box-shadow:inLeft`  `box-shadow:in_center`

![图片阴影入口](https://github.com/qrpcode/wordgo/blob/master/api/textapi.assets/%E5%9B%BE%E7%89%87%E9%98%B4%E5%BD%B1%E5%85%A5%E5%8F%A3.png?raw=true)



#### reflection 映像

只支持office默认的一些样式，对应属性关系和样式请参考下图

![映像](https://github.com/qrpcode/wordgo/blob/master/api/textapi.assets/%E6%98%A0%E5%83%8F.png?raw=true)

中间间隔符号使用 "-" 或 "_" 或 不写都可以正常解析。不区分大小写。

示例：`reflection:near`  `reflection:far-max`

![映像入口](https://github.com/qrpcode/wordgo/blob/master/api/textapi.assets/%E6%98%A0%E5%83%8F%E5%85%A5%E5%8F%A3.png?raw=true)



#### border / border-style / border-width / border-color 边框样式

border书写顺序为：宽度  样式  颜色

宽度单位为磅，支持小数，写px或者pt或者不写都可以被正常解析

样式支持：

> "solid", "dotted", "sysDash", "dashed", "dashDot", "lgDash", "lgDashDot", "lgDashDotDot"

颜色：16位颜色，带不带 # 都行

border-style / border-width / border-color 支持单独定义

示例：`border:1.5 solid #FF0000 `  `border-width:1pt`

![图片边框](https://github.com/qrpcode/wordgo/blob/master/api/textapi.assets/%E5%9B%BE%E7%89%87%E8%BE%B9%E6%A1%86.png?raw=true)



#### left / top 左/上偏移量

只有在position 属性为 fixed 或 absolute时有效，可以查看上面position 属性图，left 和 top是距离锚点的位置。

支持页面百分比，但是必须有 % ；也支持以磅为单位，pt、px结尾或者不写都会被解析为磅



#### soft-edge 柔化边缘

柔化边缘单位为磅，写px或者pt或者不写都可以被正常解析

示例：`soft-edge:20` `soft-edge:15pt`

![柔化边缘](https://github.com/qrpcode/wordgo/blob/master/api/textapi.assets/%E6%9F%94%E5%8C%96%E8%BE%B9%E7%BC%98.png?raw=true)



#### margin / margin-top / margin-left / margin-bottom / margin-bottom 图像外边距

示例（和html中略有不同）：

`margin:10 5 15 20`  上外边距是 10pt；右外边距是 5pt；下外边距是 15pt；左外边距是 20pt

`margin:10pt 5pt 15pt;`  上外边距是 10pt；右边距是 5p；下外边距是 15px

`margin:10px 5px;` 上外边距是 10pt；右外边距是 5pt

`margin:10px;`  上边距是10pt

> 这里因为担心pt不常用大家写错，px也会被当成pt解析（px无法直接转换pt）

![边距](https://github.com/qrpcode/wordgo/blob/master/api/textapi.assets/%E8%BE%B9%E8%B7%9D.png?raw=true)



### 常见问题

> Q：如果可以写pt也可以写px的地方同样数字单位不同效果有影响吗？
>
> A：没影响。只是在网页经常用px，在word中只有磅为单位，所以直接不管写啥都算对了。但是我们建议使用pt，只有图片大小使用px

> Q：为什么我设置的margin不起作用？
>
> A：设置了阴影属性、倒影属性等都会自动进行调整合适位置，可能导致原来margin失效