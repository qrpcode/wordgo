package com.github.qrpcode.util;

import com.github.qrpcode.config.Css;
import com.github.qrpcode.config.Docx;

import java.util.Map;
import java.util.concurrent.locks.Lock;
import java.util.concurrent.locks.ReentrantLock;

/**
 * css样式解析转换工具类
 * @author QiRuipeng
 */
public class TextUtil implements Runnable{

    private StringBuilder xml;
    private Map<String, String> map;

    private Lock mapLock = new ReentrantLock();
    private Lock xmlLock = new ReentrantLock();


    public StringBuilder getXml() {
        return xml;
    }

    public Map<String, String> getMap() {
        return map;
    }

    public TextUtil(StringBuilder xml, Map<String, String> map) {
        this.xml = xml;
        this.map = map;
    }

    /**
     * 用于解析当前css键值对，解析后会将希望用作替换的文本和被替换文本传回进行替换
     * 因为是多线程，所以只能保证错误值不被解析，多个正确值没有办法指定最后一个解析成功
     */
    public void run() {
        String styleKey = "";
        String styleValue = "";
        String oldString = "";
        String newString = "";

        mapLock.lock();
        try{
            String[] strings = StrUtil.getMapAndRemove(map);
            styleKey = strings[0];
            styleValue = strings[1];
        }catch (Exception e){
            return;
        }finally {
            mapLock.unlock();
        }

        //字体属性
        if(Css.FONT_FAMILY.equals(styleKey)){
            oldString = Css.A_FONT_FAMILY_A;
            newString = "<w:rFonts w:ascii=\"" + styleValue + "\" w:eastAsia=\"" + styleValue + "\" w:hAnsi=\"" +
                    styleValue + "\" w:hint=\"eastAsia\"/>";
        }
        //字号大小
        if(Css.FONT_SIZE.equals(styleKey)){
            int fontSize = StrUtil.fontSizeToInt(styleValue);
            oldString = Css.A_FONT_SIZE_A;
            newString = "<w:sz w:val=\"" + fontSize + "\"/><w:szCs w:val=\"" + fontSize + "\"/>";
        }
        //是否加粗
        if(Css.FONT_WEIGHT.equals(styleKey)){
            if("bold".equals(styleValue)){
                oldString = Css.A_FONT_WEIGHT_A;
                newString = "<w:b/>";
            }
        }
        //是否斜体
        if(Css.FONT_STYLE.equals(styleKey)){
            if("italic".equals(styleValue) || "oblique".equals(styleValue)){
                oldString = Css.A_FONT_STYLE_A;
                newString = "<w:i/>";
            }
        }
        //设置背景色
        if(Css.BACKGROUND_COLOR.equals(styleKey)){
            String color = StrUtil.textBackgroundToString(styleValue);
            if(color != null){
                oldString = Css.A_BACKGROUND_COLOR_A;
                newString = "<w:highlight w:val=\"" + color + "\"/>";
            }
        }
        //设置文字颜色
        if(Css.COLOR.equals(styleKey)){
            String color = StrUtil.isColor(styleValue);
            if(color != null){
                oldString = Css.A_COLOR_A;
                newString = "<w:color w:val=\"" + color + "\"/>";
            }
        }
        //设置文字下划线
        if(Css.TEXT_DECORATION.equals(styleKey)){
            String decoration = StrUtil.textDecorationToString(styleValue);
            if(decoration != null){
                oldString = Css.A_TEXT_DECORATION_A;
                newString =  decoration;
            }
        }
        //设置着重号
        if(Css.FONT_EMPHASIS.equals(styleKey)){
            if("point".equals(styleValue.toLowerCase())){
                oldString = Css.A_FONT_EMPHASIS_A;
                newString = "<w:em w:val=\"dot\"/>";
            }
        }
        //文本水平缩放
        if(Css.FONT_SCALE.equals(styleKey)){
            styleValue = styleValue.replace("%", "");
            if(StrUtil.isNumeric(styleValue)) {
                int scale = Math.abs((int)Double.parseDouble(styleValue));
                if (scale > 0 && scale < 100) {
                    oldString = Css.A_FONT_SCALE_A;
                    newString = "<w:w w:val=\"" + scale + "\"/>";
                }
            }
        }

        xmlLock.lock();
        try {
            StrUtil.replace(xml, oldString, newString);
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            xmlLock.unlock();
        }
    }

    /**
     * 如果用户有下划线颜色设定
     * @param xml 提前解析的xml
     * @param colorStr 颜色设定
     * @return 新的xml
     */
    static StringBuilder decorationColor(StringBuilder xml, String colorStr) {
        String color = StrUtil.isColor(colorStr);
        if(color != null){
            return StrUtil.replace(xml, "@text-decoration-color@", "<w:u w:val=\"@text-decoration@\" w:color=\"" + color + "\"/>");
        }else{
            return xml;
        }
    }

    /**
     * 如果没有颜色设定
     * @param xml 默认xml
     * @return 返回新的xml
     */
    static StringBuilder textDecoration(StringBuilder xml) {
        return StrUtil.replace(xml, "@text-decoration-color@", "<w:u w:val=\"@text-decoration@\" />");
    }

    static String lineHeight(String styleValue, int fontSize){
        String newString = "";
        if(styleValue.contains("%")){
            //百分比接收 25%-1000% 之间的数字
            styleValue = styleValue.replace("%", "");
            if(StrUtil.isNumeric(styleValue)){
                //如果是数字在处理
                int height = Integer.parseInt(styleValue);
                if(height >= 25 && height <= 1000){
                    height = (int)Math.round(height * 240 * 0.01);
                    newString = "<w:spacing w:line=\"" + height + "\" w:lineRule=\"auto\"/><w:sz w:val=\"" + fontSize + "\"/><w:szCs w:val=\"" + fontSize + "\"/>";
                }
            }
        }else{
            styleValue = styleValue.replace("px", "").replace("pt", "");
            if(StrUtil.isNumeric(styleValue)){
                //自定义行距，磅为单位，必须为整数支持 1 - 500
                int height = Integer.parseInt(styleValue);
                if(height >= 1 && height <= 500){
                    height = height * 20;
                    newString = "<w:spacing w:line=\"" + height + "\" w:lineRule=\"exact\"/>";
                }
            }
        }
        return newString;
    }
}
