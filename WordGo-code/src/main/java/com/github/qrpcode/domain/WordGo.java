package com.github.qrpcode.domain;


import com.github.qrpcode.config.Css;
import com.github.qrpcode.util.IoUtil;
import com.github.qrpcode.util.WordFileUtilImpl;

import java.io.File;
import java.util.Map;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;


/**
 * word文档类
 * @author QiRuipeng
 */
public class WordGo {

    private WordCore wordCore;
    private WordPaper wordPaper;

    /**
     * 临时储存一行的xml内容
     */
    private StringBuilder lineXml;
    /**
     * 临时储存一行的css内容
     */
    private StringBuilder lineCss;
    private StringBuilder wordXml;
    private String tempPath;
    /**
     * 页眉页脚
     */
    private PaperOut paperHead;
    private PaperOut paperFoot;

    public void setPaperHead(PaperOut paperHead) {
        this.paperHead = paperHead;
    }

    public void setPaperFoot(PaperOut paperFoot) {
        this.paperFoot = paperFoot;
    }

    private ExecutorService es;

    public WordGo() {
        this.wordXml = new StringBuilder();
        this.lineXml = new StringBuilder();
        this.lineCss = new StringBuilder();
        this.tempPath = new WordFileUtilImpl().startWordGo();
        this.wordPaper = new WordPaper("A4", false);
        this.paperHead = new PaperOut();
        this.paperFoot = new PaperOut();
        es = Executors.newFixedThreadPool(6);
    }

    public void setWordCore(WordCore wordCore) {
        this.wordCore = wordCore;
    }

    public WordGo(WordCore wordCore) {
        this();
        setWordCore(wordCore);
    }

    public WordGo(String paperType) {
        this();
        this.wordPaper = new WordPaper(paperType, false);
    }

    public WordGo(WordPaper wordPaper) {
        this();
        this.wordPaper = wordPaper;
    }

    public WordGo(WordCore wordCore, WordPaper wordPaper) {
        this();
        setWordCore(wordCore);
        this.wordPaper = wordPaper;
    }

    public void addLine(String text, String css){
        add(text + "\\n", css);
    }

    public void add(String text, String css){
        text = text.replaceAll("(\\r\\n|\\n|\\n\\r)","@@@[br]@@@");
        String[] texts = text.split("@@@");
        for (String s : texts) {
            if(s != null && s.length() > 0 && !"[br]".equals(s)){
                //行内文本处理
                lineXml.append(new WordFileUtilImpl().wordTestDo(s, css, es));
                lineCss.append(css);
            }else if(s != null && s.length() > 0){
                //行文本合成
                wordXml.append(new WordFileUtilImpl().wordLineDo(lineXml, lineCss));
                lineXml = new StringBuilder();
                lineCss = new StringBuilder();
            }
        }
    }

    /**
     * 更换到下一个页面中，不管当前页面写了多少，都切换到下一页
     */
    public void newPage(){
        addLastLine();
        wordXml.append(new WordFileUtilImpl().wordNewPageDo());
    }

    public void addImg(String uri, String css){
        Map<String, String> cssMap = new WordFileUtilImpl().imgCssDo(css);
        //经过清洗文档定位或者锚点定位一定是
        if(Css.TRUE.equals(cssMap.get(Css.NEW_LINE))){
            //如果图片单独显示在一行，先把上一行的给解析完，留出地方储存这一行
            addLastLine();
        }
        //文档定位：直接插到文档开头
        if(Css.POSITION_FIXED.equals(cssMap.get(Css.POSITION))){
            wordXml.insert(0, new WordFileUtilImpl().wordImgDo(uri, cssMap, wordPaper.getPaperWidth(), wordPaper.getPaddingLeft(),
                    wordPaper.getPaddingRight(), wordPaper.getPaperHeight(), wordPaper.getPaperHeader(), wordPaper.getPaddingBottom(),
                    tempPath, es));
        }
        //锚点定位：直接插入到文档当前最后
        if(Css.POSITION_ABSOLUTE.equals(cssMap.get(Css.POSITION))){
            lineXml.insert(0, new WordFileUtilImpl().wordImgDo(uri, cssMap, wordPaper.getPaperWidth(), wordPaper.getPaddingLeft(),
                    wordPaper.getPaddingRight(), wordPaper.getPaperHeight(), wordPaper.getPaperHeader(), wordPaper.getPaddingBottom(),
                    tempPath, es));
            lineCss.append(css);
            //锚点定位用户是否单独一行其实没有影响
        }
        //流定位：直接插入行最后
        if(Css.POSITION_STATIC.equals(cssMap.get(Css.POSITION))){
            String lineFlag = cssMap.get(Css.NEW_LINE);
            lineXml.append(new WordFileUtilImpl().wordImgDo(uri, cssMap, wordPaper.getPaperWidth(), wordPaper.getPaddingLeft(),
                    wordPaper.getPaddingRight(), wordPaper.getPaperHeight(), wordPaper.getPaperHeader(), wordPaper.getPaddingBottom()
                    , tempPath, es));
            lineCss.append(css);
            //如果希望单独一行，则解析单独一行
            if(Css.TRUE.equals(lineFlag)){
                add("\n","");
            }
        }
    }

    /**
     * 添加表格
     * @param wordTable 表格对象
     */
    public void addTable(WordTable wordTable){
        add(wordTable);
    }

    /**
     * 添加表格
     * @param wordTable 表格对象
     */
    public void add(WordTable wordTable) {
        addLastLine();
        wordXml.append(new WordFileUtilImpl().wordTableDo(wordTable, tempPath, es, wordPaper.getPaperHeight(), wordPaper.getPaddingTop(),
                wordPaper.getPaddingBottom(), wordPaper.getPaperHeader(), wordPaper.getPaperFooter(), wordPaper.getPaperWidth(),
                wordPaper.getPaddingLeft(), wordPaper.getPaddingRight()));
    }

    public void addHead(PaperOut header){
        this.paperHead = header;
    }
    public void addFoot(PaperOut footer){
        this.paperFoot = footer;
    }

    /**
     * 最终生成一个docx
     */
    public void create(String fileWay){
        //添加最后一行数据
        addLastLine();
        //生成页眉页脚信息
        StringBuilder paperOut = new StringBuilder();
        boolean paperHeadGet = (paperHead != null && !"".equals(paperHead.getContext().toString()));
        boolean paperFootGet = (paperFoot != null && !"".equals(paperFoot.getContext().toString()));
        if(paperHeadGet || paperFootGet){
            paperOut = new WordFileUtilImpl().wordPaperOutDo(tempPath, paperHead.getText(), paperHead.getCss(), paperHead.getImg(),
                    paperHead.getImgCss(), paperHead.getContext(), paperHead.getContextCss(),
                    paperFoot.getText(), paperFoot.getCss(), paperFoot.getImg(), paperFoot.getImgCss(), paperFoot.getContext(),
                    paperFoot.getContextCss(), wordPaper.getPaperWidth(), wordPaper.getPaddingLeft(), wordPaper.getPaddingRight(),
                    wordPaper.getPaperHeight(), wordPaper.getPaperHeader(), wordPaper.getPaddingBottom(), es);
        }
        //生成纸张类型信息
        wordXml.append(new WordFileUtilImpl().wordPaperDo(wordPaper.getPaperWidth(), wordPaper.getPaperHeight(), wordPaper.getPaddingTop(),
                wordPaper.getPaddingRight(), wordPaper.getPaddingBottom(), wordPaper.getPaddingLeft(), wordPaper.getPaperHeader(),
                wordPaper.getPaperFooter(), wordPaper.getCols(), wordPaper.isLandscape(), paperOut.toString()));
        //生成整个文档
        new WordFileUtilImpl().wordCreate(wordCore, tempPath, wordXml, fileWay);
        es.shutdown();
    }

    /**
     * 判断最后一行是不是有数据，有的话就加上一行
     */
    private void addLastLine(){
        if(lineXml != null && !"".equals(lineXml.toString())){
            add("\n","");
        }
    }
}
