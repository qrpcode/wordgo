package com.github.qrpcode.domain;

import com.github.qrpcode.config.Docx;


/**
 * 管理文档纸张的类
 * @author QiRuipeng
 */
public class WordPaper {
    private int paperWidth;
    private int paperHeight;
    private int paddingTop;
    private int paddingRight;
    private int paddingBottom;
    private int paddingLeft;
    private int paperHeader;
    private int paperFooter;
    private boolean landscape;
    private StringBuilder cols;

    public WordPaper(String paper){
        this(paper, false);
    }

    public WordPaper(String paper, boolean landscape) {
        paper = paper.replace(" ", "").replace("-", "").replace("_", "");
        if(Docx.DOCX_PAPER_A4.equals(paper.toLowerCase())){
            this.paperWidth = 11906;
            this.paperHeight = 16838;
        }else if(Docx.DOCX_PAPER_A3.equals(paper.toLowerCase())){
            this.paperWidth = 16838;
            this.paperHeight = 23811;
        }else if(Docx.DOCX_PAPER_A5.equals(paper.toLowerCase())){
            this.paperWidth = 8391;
            this.paperHeight = 11906;
        }else if(Docx.DOCX_PAPER_B4_JIS.equals(paper.toLowerCase())){
            this.paperWidth = 14570;
            this.paperHeight = 20636;
        }else if(Docx.DOCX_PAPER_B5_JIS.equals(paper.toLowerCase())){
            this.paperWidth = 10318;
            this.paperHeight = 14570;
        }else if(Docx.DOCX_PAPER_TABLOID.equals(paper.toLowerCase())){
            this.paperWidth = 15840;
            this.paperHeight = 24480;
        }else if(Docx.DOCX_PAPER_STATEMENT.equals(paper.toLowerCase())){
            this.paperWidth = 7920;
            this.paperHeight = 12240;
        }else if(Docx.DOCX_PAPER_EXECUTIVE.equals(paper.toLowerCase())){
            this.paperWidth = 10440;
            this.paperHeight = 15120;
        }else if(Docx.DOCX_PAPER_LETTER.equals(paper.toLowerCase())){
            this.paperWidth = 12240;
            this.paperHeight = 15840;
        }else if(Docx.DOCX_PAPER_LAW.equals(paper.toLowerCase())){
            this.paperWidth = 12240;
            this.paperHeight = 20160;
        }else if(Docx.DOCX_PAPER_16CUT.equals(paper.toLowerCase())){
            this.paperWidth = 10433;
            this.paperHeight = 14742;
        }else if(Docx.DOCX_PAPER_32CUT.equals(paper.toLowerCase())){
            this.paperWidth = 7371;
            this.paperHeight = 10433;
        }else if(Docx.DOCX_PAPER_BIG32CUT.equals(paper.toLowerCase())){
            this.paperWidth = 7938;
            this.paperHeight = 11510;
        }else{
            throw new RuntimeException("您定义的纸张类型没有预设值，请通过自定义方式传递需要的纸张大小和页边距！There is no preset value for the paper type you set. Please pass the required paper size and margins in a custom way！");
        }
        if(this.paperWidth > 0){
            this.paddingTop = 1440;
            this.paddingRight = 1800;
            this.paddingBottom = 1440;
            this.paddingLeft = 1800;
            this.paperHeader = 851;
            this.paperFooter = 992;
            this.landscape = landscape;
            cols = new StringBuilder("<w:cols w:space=\"425\"/>");
        }
    }

    public WordPaper(int paperWidth, int paperHeight, boolean landscape){
        this.paperWidth = (int)(paperWidth * Docx.DOCX_PAPER);
        this.paperHeight = (int)(paperHeight * Docx.DOCX_PAPER);
        this.paddingTop = 1440;
        this.paddingRight = 1800;
        this.paddingBottom = 1440;
        this.paddingLeft = 1800;
        this.paperHeader = 851;
        this.paperFooter = 992;
        this.landscape = landscape;
        cols = new StringBuilder("<w:cols w:space=\"425\"/>");
    }

    public WordPaper(int paperWidth, int paperHeight, int paddingTop, int paddingRight, int paddingBottom, int paddingLeft,
                     int paperHeader, int paperFooter, boolean landscape) {
        this.paperWidth = (int)(paperWidth * Docx.DOCX_PAPER);
        this.paperHeight = (int)(paperHeight * Docx.DOCX_PAPER);
        this.paddingTop = (int)(paddingTop * Docx.DOCX_PAPER);
        this.paddingRight = (int)(paddingRight * Docx.DOCX_PAPER);
        this.paddingBottom = (int)(paddingBottom * Docx.DOCX_PAPER);
        this.paddingLeft = (int)(paddingLeft * Docx.DOCX_PAPER);
        this.paperHeader = (int)(paperHeader * Docx.DOCX_PAPER);
        this.paperFooter = (int)(paperFooter * Docx.DOCX_PAPER);
        this.landscape = landscape;
        cols = new StringBuilder("<w:cols w:space=\"425\"/>");
    }

    /**
     * 设定分栏数量和间隔
     * 设置失败可能：分栏是负数或0、间距之和已经大于了当前纸张总大小
     * @param num 需要的分栏数量
     * @param space 每两个分栏之间间隔的字符数，推荐值2.02
     * @param sep 是否需要分栏线
     * @return 返回是否设置成功
     */
    public boolean colsAverage(int num, double space, boolean sep){
        if(num < 1){
            return false;
        }
        int reallySpace = (int)Math.ceil(space * 210);
        if(reallySpace > 0 && reallySpace * num < paperWidth - paddingLeft - paddingRight - num){
            if(num == 1){
                cols = new StringBuilder("<w:cols w:space=\"425\"/>");
            }else{
                String spaceString = "";
                if(sep){
                    spaceString = "w:sep=\"1\" ";
                }
                cols = new StringBuilder("<w:num=\"" + num + "\" " + spaceString + reallySpace + "\"/>");
            }
            return true;
        }else{
            return false;
        }
    }

    int getPaperWidth() {
        return paperWidth;
    }

    int getPaperHeight() {
        return paperHeight;
    }

    int getPaddingTop() {
        return paddingTop;
    }

    int getPaddingRight() {
        return paddingRight;
    }

    int getPaddingBottom() {
        return paddingBottom;
    }

    int getPaddingLeft() {
        return paddingLeft;
    }

    int getPaperHeader() {
        return paperHeader;
    }

    StringBuilder getCols() {
        return cols;
    }

    boolean isLandscape() {
        return landscape;
    }

    int getPaperFooter() {
        return paperFooter;
    }

    public void setPaperWidth(int paperWidth) {
        this.paperWidth = paperWidth;
    }

    public void setPaperHeight(int paperHeight) {
        this.paperHeight = paperHeight;
    }

    public void setPaddingTop(int paddingTop) {
        this.paddingTop = paddingTop;
    }

    public void setPaddingRight(int paddingRight) {
        this.paddingRight = paddingRight;
    }

    public void setPaddingBottom(int paddingBottom) {
        this.paddingBottom = paddingBottom;
    }

    public void setPaddingLeft(int paddingLeft) {
        this.paddingLeft = paddingLeft;
    }

    public void setPaperHeader(int paperHeader) {
        this.paperHeader = paperHeader;
    }

    public void setPaperFooter(int paperFooter) {
        this.paperFooter = paperFooter;
    }
}
