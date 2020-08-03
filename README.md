# Word GO

WordGO - Making it easier for Java to generate word documents

【choose language：中文介绍】

### Background

In traditional Java Word generation, template files are usually created manually and then imported. If you don't want to create a template and want less code, it's a good idea to choose Word Go~~

### Installation

#### Manually import the jar package

* IDEA import: Click File-Project Structure; then find Modules on the left and click; finally, click the green + sign on the right, select JARs or directories and select the jar package to be imported.
* Eclipse import: right-click "Project" → select Properties, select Java Build Path in the list on the left side of the pop-up dialog box

#### maven import

```xml
<dependency>
 <groupId>com.github.qrpcode</groupId>
 <artifactId>wordgo</artifactId>
 <version>1.0-SNAPSHOT</version>
</dependency>
```

### Environment dependency and compatibility

It can be used as long as it can run java, it does not rely on any third-party Office applications

See the table for compatibility:

![我的兼容性](https://github.com/qrpcode/wordgo/blob/master/api/textapi.assets/enjianrong.png?raw=true)

### Use

Let’s create a "Hello World" first

```java
WordGo wordGo = new WordGo();
//Create a new word
wordGo.add("Hello World", "font-size: 15; color: #FF0000");
//Fill in the data to view the corresponding function description
wordGo.create("C:\\demo.docx");
//Finally, it can be generated. The parameter is the generation directory, which must have a file name and end with. Docx
```

Yes, it is very similar to Css, it is easy to use~~

Code support JDK1.5 + (inclusive)

#### Corresponding function description

####  (can be learned in 10 minutes)

👉 [WordGO Construct and generate files](https://github.com/qrpcode/wordgo/blob/master/api/wordgoapi.md)

👉 [Set document attribute information](https://github.com/qrpcode/wordgo/blob/master/api/coreapi.md)

👉 [Set paper size and margins](https://github.com/qrpcode/wordgo/blob/master/api/paperapi.md)

👉 [Text, page change related operations](https://github.com/qrpcode/wordgo/blob/master/api/textapi.md)

👉 [Create, fill, and add tables related operations](https://github.com/qrpcode/wordgo/blob/master/api/tableapi.md)

👉 [Picture related operations](https://github.com/qrpcode/wordgo/blob/master/api/imgapi.md)

👉 [Header and footer related operations](https://github.com/qrpcode/wordgo/blob/master/api/paperoutapi.md)

### Main project leader

[@qrpcode](https://github.com/qrpcode)

### Join

Yes, I also think the code written by me is ~~a bit~~ (very) messy

Come and help me, just pull request after Fork~

### Open source agreement

[Apache-2.0 License](https://github.com/qrpcode/wordgo/blob/master/LICENSE)

(In other words, it can be commercialized, please see the agreement in detail~~)



### 💖 If you find it useful, remember to star ⭐



> PS：Welcome to download and try it out. If you need to apply it to commercial projects, please be sure to **test all possible functions in advance** to avoid **generation deviation**.
>
> Feel free to leave a message if you find a bug!