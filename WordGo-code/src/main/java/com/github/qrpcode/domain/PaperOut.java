package com.github.qrpcode.domain;

import com.github.qrpcode.config.Docx;

import java.util.ArrayList;
import java.util.List;

/**
 * 页眉页脚
 * @author Qiruipeng
 */
public class PaperOut {
    private StringBuilder text;
    private StringBuilder css;
    private List<String> img;
    private List<String> imgCss;
    private List<String> context;
    private List<String> contextCss;

    /**
     * 这些get方法wordGo能获得就行，外面的获得就很诡异，所以不能public
     * @return 返回
     */
    StringBuilder getText() {
        return text;
    }

    StringBuilder getCss() {
        return css;
    }

    List<String> getImg() {
        return img;
    }

    List<String> getImgCss() {
        return imgCss;
    }

    List<String> getContext() {
        return context;
    }

    List<String> getContextCss() {
        return contextCss;
    }

    public PaperOut(){
        text = new StringBuilder();
        css = new StringBuilder();
        img = new ArrayList<String>();
        imgCss = new ArrayList<String>();
        context = new ArrayList<String>();
        contextCss = new ArrayList<String>();
    }

    public PaperOut(String paperText){
        this();
        addText(paperText);
    }

    public PaperOut(String paperText, String paperCss){
        this();
        addText(paperText);
        css = new StringBuilder(paperCss);
    }

    public void addText(String paperText){
        text = addTextTo(text, paperText.replace("#", ""));
    }

    public void addCss(String paperCss){
        css = addTextTo(css, paperCss);
    }

    public void addStyleText(String styleText, String styleCss){
        if(styleText != null && styleCss != null && !"".equals(styleText) && !"".equals(styleCss)){
            context.add(styleText);
            contextCss.add(styleCss);
            addTextTo(text, Docx.DOCX_TABLE_CSS_KEY + (context.size() - 1));
        }
    }

    public void addImg(String imgUri, String imgStyleCss){
        if(imgUri != null && imgStyleCss != null && !"".equals(imgUri) && !"".equals(imgStyleCss)){
            img.add(imgUri);
            imgCss.add(imgStyleCss);
            addTextTo(text, Docx.DOCX_TABLE_IMG_KEY + (img.size() - 1));
        }
    }

    /**
     * 添加页码，默认样式（就一个数字那种 ~）
     */
    public void addPage(){
        addPage(1, "");
    }

    /**
     * 添加页码
     * @param template 样式模板，可以接收
     */
    public boolean addPage(int template, String numStyleCss){
        if(template < 1 || template > 3){
            return false;
        }else{
            context.add(Docx.DOCX_PAGE);
            contextCss.add(numStyleCss + "template:" + template);
            addTextTo(text, Docx.DOCX_PAGE + (context.size() - 1));
        }
        return false;
    }

    private StringBuilder addTextTo(StringBuilder sb, String text){
        if(sb == null || "".equals(sb.toString())){
            sb = new StringBuilder(text);
        }else{
            sb.append(Docx.DOCX_TABLE_HASH);
            sb.append(text);
        }
        return sb;
    }

}
