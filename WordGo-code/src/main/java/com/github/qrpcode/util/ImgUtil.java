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
public class ImgUtil implements Runnable{

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

    public ImgUtil(StringBuilder xml, Map<String, String> map) {
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

        //宽高、阴影、边框可以直接解析
        if(Css.WIDTH.equals(styleKey) || Css.HEIGHT.equals(styleKey) || Css.BOX_SHADOW.equals(styleKey) || Css.REFLECTION.equals(styleKey) ||
                Css.BORDER.equals(styleKey) || Css.MARGIN_TOP.equals(styleKey) || Css.MARGIN_RIGHT.equals(styleKey) ||
                Css.MARGIN_BOTTOM.equals(styleKey) || Css.MARGIN_LEFT.equals(styleKey)) {
            oldString = Css.AT + styleKey + Css.AT;
            newString = styleValue;
        }
        //偏移量
        if(Css.LEFT.equals(styleKey) || Css.TOP.equals(styleKey)){
            oldString = Css.AT + styleKey + Css.AT;
            newString = styleValue;
        }
        //柔化边缘
        if(Css.SOFT_EDGE.equals(styleKey)){
            if(StrUtil.isNumeric(styleValue)){
                //其实这里单位是磅，为了防止用户后面写px（像素）或者 pt（磅）报错，我们都给他替换掉。px和pt没法转化，所以不能转换！
                int soft = (int)Double.parseDouble(styleValue.replace("px", "").replace("pt", ""));
                if(soft > 0 && soft <= 100){
                    oldString = Css.A_SOFT_EDGE_A;
                    newString = "<a:softEdge rad=\"" + Css.SOFT_EDGE_TIMES * soft + "\"/>";
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


}
