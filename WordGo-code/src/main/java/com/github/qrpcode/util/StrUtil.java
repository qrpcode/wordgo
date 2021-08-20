package com.github.qrpcode.util;

import com.github.qrpcode.config.Css;
import com.github.qrpcode.config.Docx;
import com.github.qrpcode.domain.WordCore;

import java.io.File;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.concurrent.locks.Lock;
import java.util.concurrent.locks.ReentrantLock;
import java.util.regex.Pattern;

/**
 * 文字处理工具类
 * @author QiRuiPeng
 */
class StrUtil {

    private static final String[] FONT = {"初号", "小初", "一号", "小一", "二号", "小二", "三号", "小三", "四号", "小四", "五号",
            "小五", "六号", "小六", "七号", "八号"};
    private static final String[] FONT_BACK = {"yellow", "green", "cyan", "magenta", "blue", "red", "darkBlue","darkCyan",
            "darkGreen", "darkMagenta", "darkRed", "darkYellow", "darkGray", "lightGray", "black"};
    private static final String[] FONT_LINE = {"single", "double", "thick", "dotted", "dash", "dotDash", "dotDotDash","wave"};
    private static final String[] STYLE_LABEL = {"@text-align@", "@font-family@", "@font-size@", "@font-weight@", "@font-style@",
            "@background-color@", "@color@", "@text-decoration@", "@font-emphasis@", "@font-scale@", "@line-height@",
            "@text-decoration-color@"};
    private static final String[] STYLE_IMG_LABEL = {"@left@", "@top@", "@height@", "@width@", "@margin-left@", "@margin-top@",
            "@margin-right@", "@margin-bottom@", "@title@", "@box-shadow@", "@reflection@", "@glow@", "@border@", "@rotate@",
            "@soft-edge@"};

    /**
     * 将css样式文本切割为Map集合
     * @param css css文本
     * @return map集合
     */
    static Map<String, String> cssToMap(String css){
        Map<String, String> map = new HashMap<String, String>();
        String[] cssList = css.split(";");
        for (String s : cssList) {
            String[] style = s.split(":");
            if(style.length == 2 && style[0] != null && !"".equals(style[0]) && style[1] != null && !"".equals(style[1])){
                map.put(style[0].trim(), style[1].trim());
            }
        }
        return map;
    }

    /**
     * StringBuilder 替换文本
     * @param sb 原文本
     * @param oldStr 被替换的文本
     * @param newStr 用作替换的文本
     * @return 替换完成的字符串
     */
    static StringBuilder replace(StringBuilder sb, String oldStr, String newStr){
        //因为css样式解析是多线程的，不加锁经常提醒索引越界，一个字符串已经改了文字导致现在文字变化直接报错，概率较小，可以删除运行几次试一试
        //如果每个需要多线程解析的css都只用一个，不加锁问题不大
        //需要多线程解析的css：有些css样式会提前解析，因为他们嵌套了其他标签
        Lock l = new ReentrantLock();
        l.lock();
//        防止 & 死循环
        int oldIndex = 0;
        try {
            while (oldStr != null && !"".equals(oldStr) && sb.indexOf(oldStr, oldIndex) > -1) {
                int oldIndexTemp = sb.indexOf(oldStr, oldIndex);
                sb.replace(sb.indexOf(oldStr, oldIndex), sb.indexOf(oldStr, oldIndex) + oldStr.length(), newStr);
                oldIndex =  oldIndexTemp + 1;
            }
        }catch (Exception e){
            //e.printStackTrace();
        }finally {
            l.unlock();
        }
        return sb;
    }

    /**
     * StringBuilder 替换文本,只替换一次，不管存在多少，不存在不替换
     * @param sb 原文本
     * @param oldStr 被替换的文本
     * @param newStr 用作替换的文本
     */
    static void replaceOnce(StringBuilder sb, String oldStr, String newStr){
        if(oldStr != null && !"".equals(oldStr) && sb.indexOf(oldStr) > -1) {
            sb.replace(sb.indexOf(oldStr), sb.indexOf(oldStr) + oldStr.length(), newStr);
        }
    }

    /**
     * 将字号文字转为数字
     * @param size 可能是数字或者是如“初号”  “一号” 这种文字
     * @return 返回文档中实际记录的数值
     */
    static int fontSizeToInt(String size){
        if(FONT[0].equals(size)){
            return 84;
        }else if(FONT[1].equals(size)){
            return 72;
        }else if(FONT[2].equals(size)){
            return 52;
        }else if(FONT[3].equals(size)){
            return 48;
        }else if(FONT[4].equals(size)){
            return 44;
        }else if(FONT[5].equals(size)){
            return 36;
        }else if(FONT[6].equals(size)){
            return 32;
        }else if(FONT[7].equals(size)){
            return 30;
        }else if(FONT[8].equals(size)){
            return 28;
        }else if(FONT[9].equals(size)){
            return 24;
        }else if(FONT[10].equals(size)){
            return 21;
        }else if(FONT[11].equals(size)){
            return 18;
        }else if(FONT[12].equals(size)){
            return 15;
        }else if(FONT[13].equals(size)){
            return 13;
        }else if(FONT[14].equals(size)){
            return 11;
        }else if(FONT[15].equals(size)){
            return 10;
        }else{
            if(isNumeric(size)){
                return Math.ceil(Double.parseDouble(size) * 2) > 0 ? (int)Math.ceil(Double.parseDouble(size) * 2) : 21;
            }else {
                //解析不了返回五号字
                return 21;
            }
        }
    }

    /**
     * 判断是不是数字类型
     * @param str 文本
     * @return 返回是不是，是数字返回true，不是返回false
     */
    static boolean isNumeric(final String str) {
        // null or empty
        if (str == null || str.length() == 0) {
            return false;
        }
        try {
            Integer.parseInt(str);
            return true;
        } catch (NumberFormatException e) {
            try {
                Double.parseDouble(str);
                return true;
            } catch (NumberFormatException ex) {
                try {
                    Float.parseFloat(str);
                    return true;
                } catch (NumberFormatException exx) {
                    return false;
                }
            }
        }
    }

    /**
     * 替换文本中的几个特殊符号
     * 无需添加其他文本（如：空格等），因为word转义字符只有这几个
     * @param text 需要转换的文本
     * @return 转换后的结果
     */
    static StringBuilder textChoose(String text){
        StringBuilder context = new StringBuilder(text);
//        防止 & 死循环
        if(context.indexOf(Docx.DOCX_AMP) > -1){
            replace(context, Docx.DOCX_AMP, Docx.DOCX_AMP_CHOOSE);
        }
        if(context.indexOf(Docx.DOCX_LT) > -1){
            replace(context, Docx.DOCX_LT, Docx.DOCX_LT_CHOOSE);
        }
        if(context.indexOf(Docx.DOCX_GT) > -1){
            replace(context, Docx.DOCX_GT, Docx.DOCX_GT_CHOOSE);
        }
        return context;
    }

    /**
     * 添加文字底色，注意只允许几个固定的颜色而且需要小写，这是一个效验函数
     * @param colorStr 传入用户输入的信息
     * @return 返回解析后的信息
     */
    static String textBackgroundToString(String colorStr){
        //用户大小写随意，需要解析正确写法，所以不能用Arrays.binarySearch
        for (String color : FONT_BACK) {
            if(color.equalsIgnoreCase(colorStr)){
                return color;
            }
        }
        return null;
    }

    /**
     * 效验下划线类型是否是合法值
     * @param decoration 用户输入的下划线单词
     * @return 正规写法下划线单词
     */
    static String textDecorationToString(String decoration){
        //用户大小写随意，需要解析正确写法，所以不能用Arrays.binarySearch
        for (String line : FONT_LINE) {
            if(line.equalsIgnoreCase(decoration)){
                return line;
            }
        }
        return null;
    }

    /**
     * 效验一个字符串是不是十六进制颜色形式
     * @param color 传入颜色信息
     * @return 返回是否可以匹配
     */
    static String isColor(String color){
        if(color != null){
            color = color.replace("#", "").toUpperCase();
            return Pattern.matches("^[0-9A-F]{6}$", color) ? color : null;
        }else{
            return null;
        }
    }

    /**
     * 清除未匹配的标记符号（ @xxx@ ）
     * @param sb 用户开始输入的文字
     */
    static void cleanNoMarry(StringBuilder sb){
        for (String label : STYLE_LABEL) {
            while(sb.indexOf(label) > -1){
                replace(sb, label, "");
            }
        }
    }

    /**
     * 对图片css进行清洗，主要是补全一些样式同时根据逻辑预处理一些多线程处理会出错的数据
     * @param css 原来的css文本
     * @return 返回的是处理好的Css map集合
     */
    static Map<String, String> imgCssToMap(String css){
        Map<String, String> cssMap = cssToMap(css);
        //如果没设置定位方式为默认流定位
        if(!cssMap.containsKey(Css.POSITION)){
            cssMap.put(Css.POSITION, Css.POSITION_STATIC);
        }
        //如果用户选择了相对于文档定位或者相对于锚点定位必须设置独立一行
        if(Css.POSITION_FIXED.equals(cssMap.get(Css.POSITION)) || Css.POSITION_ABSOLUTE.equals(cssMap.get(Css.POSITION))){
            cssMap.put(Css.NEW_LINE, Css.TRUE);
        }else if(!cssMap.containsKey(Css.NEW_LINE)){
            cssMap.put(Css.NEW_LINE, Css.FALSE);
        }
        return cssMap;
    }

    /**
     * 针对绝对定位和相对定位进行数据的整理工作
     * @param cssMap css的map
     */
    static void positionCssClean(Map<String, String> cssMap, int paperWidth, int paddingLeft, int paddingRight,
                                 int paperHeight, int paperHead, int paperBottom){
        //首先确定不是流定位
        if(Css.POSITION_FIXED.equals(cssMap.get(Css.POSITION)) || Css.POSITION_ABSOLUTE.equals(cssMap.get(Css.POSITION))){
            if(!cssMap.containsKey(Css.LEFT)){
                cssMap.put(Css.LEFT, "0");
            }
            if(!cssMap.containsKey(Css.TOP)){
                cssMap.put(Css.TOP, "0");
            }
            //确定是不是以百分数进行的定位，如果包含百分比
            if(cssMap.get(Css.LEFT).contains(Css.PERCENT)){
                String leftNum = cssMap.get(Css.LEFT).replace(Css.PERCENT, "");
                if(isNumeric(leftNum)) {
                    cssMap.put(Css.LEFT,  (int)(Double.parseDouble(leftNum) * 0.01 * Docx.DOCX_PROPORTION *
                            (paperWidth - paddingLeft - paddingRight)) + "");
                }else {
                    cssMap.put(Css.LEFT, "0");
                }
            }else{
                cssMap.put(Css.LEFT,cssMap.get(Css.LEFT).replace(Css.PX, "").replace(Css.PT, ""));
                cssMap.put(Css.LEFT, sizeNumDo(cssMap.get(Css.LEFT)) + "");
            }
            if(cssMap.get(Css.TOP).contains(Css.PERCENT)){
                String rightNum = cssMap.get(Css.TOP).replace(Css.PERCENT, "");
                if(isNumeric(rightNum)) {
                    cssMap.put(Css.TOP,  (int)(Double.parseDouble(rightNum) * 0.01 * Docx.DOCX_PROPORTION * (paperHeight -
                            paperHead - paperBottom)) + "");
                }else {
                    cssMap.put(Css.TOP, "0");
                }
            }else{
                cssMap.put(Css.TOP,cssMap.get(Css.TOP).replace(Css.PX, "").replace(Css.PT, ""));
                cssMap.put(Css.TOP, sizeNumDo(cssMap.get(Css.TOP)) + "");
            }
        }
    }

    /**
     * 整理map集合中的width和height属性、margin属性，宽高用户不写则读取图片宽高
     * @param cssMap css的map
     * @param imgUri 图片地址，这里地址应该是已经复制到临时目录后的地址
     */
    static void sizeCssClean(Map<String, String> cssMap, String imgUri){
        //首先把margin属性拆分，map必须保证没有margin
        if(cssMap.containsKey(Css.MARGIN)){
            String[] margins = cssMap.get(Css.MARGIN).split(" ");
            for (int i = 0; i < margins.length; i++) {
                switch (i){
                    case 1:{
                        cssMap.put(Css.MARGIN_TOP, sizeNumDo(margins[i]) + "");
                        break;
                    }
                    case 2:{
                        cssMap.put(Css.MARGIN_RIGHT, sizeNumDo(margins[i]) + "");
                        break;
                    }
                    case 3:{
                        cssMap.put(Css.MARGIN_BOTTOM, sizeNumDo(margins[i]) + "");
                        break;
                    }
                    case 4:{
                        cssMap.put(Css.MARGIN_LEFT, sizeNumDo(margins[i]) + "");
                        break;
                    }
                    default:break;
                }
            }
            cssMap.remove(Css.MARGIN);
        }
        //如果没设置margin边距
        if(!cssMap.containsKey(Css.MARGIN_LEFT)){
            cssMap.put(Css.MARGIN_LEFT, "0");
        }
        if(!cssMap.containsKey(Css.MARGIN_RIGHT)){
            cssMap.put(Css.MARGIN_RIGHT, "0");
        }
        if(!cssMap.containsKey(Css.MARGIN_TOP)){
            cssMap.put(Css.MARGIN_TOP, "0");
        }
        if(!cssMap.containsKey(Css.MARGIN_BOTTOM)){
            cssMap.put(Css.MARGIN_BOTTOM, "0");
        }
        //处理 width 和 height
        File img = new File(imgUri);
        if(cssMap.containsKey(Css.WIDTH) && isNumeric(cssMap.get(Css.WIDTH).replace(Css.PX, ""))){
            cssMap.put(Css.WIDTH, Math.abs((int)Double.parseDouble(cssMap.get(Css.WIDTH).replace(Css.PX, "")) * Docx.DOCX_EUM) + "");
        }else{
            cssMap.put(Css.WIDTH, Docx.DOCX_EUM * IoUtil.getImgWidth(img) + "");
        }
        if(cssMap.containsKey(Css.HEIGHT) && isNumeric(cssMap.get(Css.HEIGHT).replace(Css.PX, ""))){
            cssMap.put(Css.HEIGHT, Math.abs((int)Double.parseDouble(cssMap.get(Css.HEIGHT).replace(Css.PX, "")) * Docx.DOCX_EUM) + "");
        }else{
            cssMap.put(Css.HEIGHT, Docx.DOCX_EUM * IoUtil.getImgHeight(img) + "");
        }
    }

    /**
     * 解析边距等数值参数
     * @param num 获取的value
     */
    private static int sizeNumDo(String num){
        String newNum = num.toLowerCase().replace(Css.PX, "").replace(Css.PT, "");
        if(StrUtil.isNumeric(newNum)) {
            return (int) Double.parseDouble(newNum) * Docx.DOCX_EUM;
        }else{
            return 0;
        }
    }

    /**
     * 获取一个文件的后缀，注意jpeg会被转换为jpg。word中只支持jpg png gif，bmp等格式会被转换后保存，如有需要储存其他格式请先转换后储存
     * @param filename 文件地址
     * @return 文件格式
     */
    static String getImgFormat(String filename){
        String format = filename.substring(filename.lastIndexOf(".")).toLowerCase().substring(1);
        if(Docx.DOCX_FORMAT_JPEG.equals(format) || Docx.DOCX_FORMAT_JPG.equals(format)){
            return Docx.DOCX_FORMAT_JPG;
        }else if(Docx.DOCX_FORMAT_PNG.equals(format) || Docx.DOCX_FORMAT_GIF.equals(format)){
            return format;
        }else{
            throw new RuntimeException("插入的不是被支持的图片格式。支持的图片格式：jpg/png/gif。The inserted image format is not supported. Supported image formats: JPG / PNG / GIF");
        }
    }

    /**
     * 获取一个字符串中，文本出现的次数
     * @param str 完整的字符串
     * @return 查询到的个数
     */
    static int getRelationCount(String str){
        int index;
        int count = 0;
        while((index = str.indexOf(Docx.DOCX_RELATIONSHIP_LABEL)) != -1){
            count++;
            str = str.substring(index + 1);
        }
        return count;
    }

    static String[] getMapAndRemove(Map<String, String> map){
        String[] str = new String[2];
        for (Map.Entry<String, String> css : map.entrySet()) {
            str[0] = css.getKey().toLowerCase();
            str[1] = css.getValue();
            map.remove(str[0]);
            break;
        }
        return str;
    }

    /**
     * 清理未使用的图片内容标签
     * @param sb 原始xml
     */
    static void cleanImgNoMarry(StringBuilder sb) {
        for (String label : STYLE_IMG_LABEL) {
            while(sb.indexOf(label) > -1){
                replace(sb, label, "");
            }
        }
    }

    /**
     * 清理图片阴影功能
     * @param cssMap css封装的map集合
     */
    static void boxShadowCssClean(Map<String, String> cssMap) {
        //获取map中的box shadow 有关信息
        if(cssMap.containsKey(Css.BOX_SHADOW)){
            //为了防止用户写错，不管是 - _ 或空格都进行替换解析
            String boxShadow = cssMap.get(Css.BOX_SHADOW).replace("-", "").replace("_", "").replace(" ", "").toLowerCase();
            //更快查找，先判断前两个字符
            int valueLength = boxShadow.length();
            boolean flag = true;

            switch (valueLength){
                case 5:{
                    if(Css.IN_TOP.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_IN_TOP);
                    }else{
                        flag = false;
                    }
                    break;
                } case 6:{
                    if(Css.OUT_TOP.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_OUT_TOP);
                        setMaxMargin(cssMap, Css.BOX_SHADOW_OUT_TOP_MARGIN_LEFT, Css.BOX_SHADOW_OUT_TOP_MARGIN_TOP, Css.BOX_SHADOW_OUT_TOP_MARGIN_RIGHT, Css.BOX_SHADOW_OUT_TOP_MARGIN_BOTTOM);
                    }else if(Css.IN_LEFT.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_IN_LEFT);
                    }else{
                        flag = false;
                    }
                    break;
                } case 7:{
                    if(Css.OUT_LEFT.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_OUT_LEFT);
                        setMaxMargin(cssMap, Css.BOX_SHADOW_OUT_LEFT_MARGIN_LEFT, Css.BOX_SHADOW_OUT_LEFT_MARGIN_TOP, Css.BOX_SHADOW_OUT_LEFT_MARGIN_RIGHT, Css.BOX_SHADOW_OUT_LEFT_MARGIN_BOTTOM);
                    }else if(Css.IN_RIGHT.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_IN_RIGHT);
                    }else{
                        flag = false;
                    }
                    break;
                } case 8:{
                    if(Css.OUT_RIGHT.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_OUT_RIGHT);
                        setMaxMargin(cssMap, Css.BOX_SHADOW_OUT_RIGHT_MARGIN_LEFT, Css.BOX_SHADOW_OUT_RIGHT_MARGIN_TOP, Css.BOX_SHADOW_OUT_RIGHT_MARGIN_RIGHT, Css.BOX_SHADOW_OUT_RIGHT_MARGIN_BOTTOM);
                    }else if(Css.IN_CENTER.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_IN_CENTER);
                    }else if(Css.IN_BOTTOM.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_IN_BOTTOM);
                    }else{
                        flag = false;
                    }
                    break;
                } case 9:{
                    if(Css.OUT_BOTTOM.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_OUT_BOTTOM);
                        setMaxMargin(cssMap, Css.BOX_SHADOW_OUT_BOTTOM_MARGIN_LEFT, Css.BOX_SHADOW_OUT_BOTTOM_MARGIN_TOP, Css.BOX_SHADOW_OUT_BOTTOM_MARGIN_RIGHT, Css.BOX_SHADOW_OUT_BOTTOM_MARGIN_BOTTOM);
                    }else if(Css.OUT_CENTER.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_OUT_CENTER);
                        setMaxMargin(cssMap, Css.BOX_SHADOW_OUT_CENTER_MARGIN_LEFT, Css.BOX_SHADOW_OUT_CENTER_MARGIN_TOP, Css.BOX_SHADOW_OUT_CENTER_MARGIN_RIGHT, Css.BOX_SHADOW_OUT_CENTER_MARGIN_BOTTOM);
                    }else if(Css.IN_TOP_LEFT.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_IN_TOP_LEFT);
                    }else{
                        flag = false;
                    }
                    break;
                } case 10:{
                    if(Css.OUT_TOP_LEFT.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_OUT_TOP_LEFT);
                        setMaxMargin(cssMap, Css.BOX_SHADOW_OUT_TOP_LEFT_MARGIN_LEFT, Css.BOX_SHADOW_OUT_TOP_LEFT_MARGIN_TOP, Css.BOX_SHADOW_OUT_TOP_LEFT_MARGIN_RIGHT, Css.BOX_SHADOW_OUT_TOP_LEFT_MARGIN_BOTTOM);
                    }else if(Css.IN_TOP_RIGHT.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_IN_TOP_RIGHT);
                    }else{
                        flag = false;
                    }
                    break;
                } case 11:{
                    if(Css.OUT_TOP_RIGHT.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_OUT_TOP_RIGHT);
                        setMaxMargin(cssMap, Css.BOX_SHADOW_OUT_TOP_RIGHT_MARGIN_LEFT, Css.BOX_SHADOW_OUT_TOP_RIGHT_MARGIN_TOP, Css.BOX_SHADOW_OUT_TOP_RIGHT_MARGIN_RIGHT, Css.BOX_SHADOW_OUT_TOP_RIGHT_MARGIN_BOTTOM);
                    }else{
                        flag = false;
                    }
                    break;
                } case 12:{
                    if(Css.IN_BOTTOM_LEFT.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_IN_BOTTOM_LEFT);
                    }else{
                        flag = false;
                    }
                    break;
                } case 13:{
                    if(Css.OUT_BOTTOM_LEFT.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_OUT_BOTTOM_LEFT);
                        setMaxMargin(cssMap, Css.BOX_SHADOW_OUT_BOTTOM_LEFT_MARGIN_LEFT, Css.BOX_SHADOW_OUT_BOTTOM_LEFT_MARGIN_TOP, Css.BOX_SHADOW_OUT_BOTTOM_LEFT_MARGIN_RIGHT, Css.BOX_SHADOW_OUT_BOTTOM_LEFT_MARGIN_BOTTOM);
                    }else if(Css.IN_BOTTOM_RIGHT.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_IN_BOTTOM_RIGHT);
                    }else{
                        flag = false;
                    }
                    break;
                } case 14:{
                    if(Css.OUT_BOTTOM_RIGHT.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_OUT_BOTTOM_RIGHT);
                        setMaxMargin(cssMap, Css.BOX_SHADOW_OUT_BOTTOM_RIGHT_MARGIN_LEFT, Css.BOX_SHADOW_OUT_BOTTOM_RIGHT_MARGIN_TOP, Css.BOX_SHADOW_OUT_BOTTOM_RIGHT_MARGIN_RIGHT, Css.BOX_SHADOW_OUT_BOTTOM_RIGHT_MARGIN_BOTTOM);
                    }else{
                        flag = false;
                    }
                    break;
                } case 17:{
                    if(Css.PERSPECTIVE_BOTTOM.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_PERSPECTIVE_BOTTOM);
                        setMaxMargin(cssMap, Css.BOX_SHADOW_PERSPECTIVE_BOTTOM_MARGIN_LEFT, Css.BOX_SHADOW_PERSPECTIVE_BOTTOM_MARGIN_TOP, Css.BOX_SHADOW_PERSPECTIVE_BOTTOM_MARGIN_RIGHT, Css.BOX_SHADOW_PERSPECTIVE_BOTTOM_MARGIN_BOTTOM);
                    }else{
                        flag = false;
                    }
                    break;
                } case 18:{
                    if(Css.PERSPECTIVE_TOP_LEFT.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_PERSPECTIVE_TOP_LEFT);
                        setMaxMargin(cssMap, Css.BOX_SHADOW_PERSPECTIVE_TOP_LEFT_MARGIN_LEFT, Css.BOX_SHADOW_PERSPECTIVE_TOP_LEFT_MARGIN_TOP, Css.BOX_SHADOW_PERSPECTIVE_TOP_LEFT_MARGIN_RIGHT, Css.BOX_SHADOW_PERSPECTIVE_TOP_LEFT_MARGIN_BOTTOM);
                    }else{
                        flag = false;
                    }
                    break;
                } case 19:{
                    if(Css.PERSPECTIVE_TOP_RIGHT.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_PERSPECTIVE_TOP_RIGHT);
                        setMaxMargin(cssMap, Css.BOX_SHADOW_PERSPECTIVE_TOP_RIGHT_MARGIN_LEFT, Css.BOX_SHADOW_PERSPECTIVE_TOP_RIGHT_MARGIN_TOP, Css.BOX_SHADOW_PERSPECTIVE_TOP_RIGHT_MARGIN_RIGHT, Css.BOX_SHADOW_PERSPECTIVE_TOP_RIGHT_MARGIN_BOTTOM);
                    }else{
                        flag = false;
                    }
                    break;
                } case 21:{
                    if(Css.PERSPECTIVE_BOTTOM_LEFT.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_PERSPECTIVE_BOTTOM_LEFT);
                        setMaxMargin(cssMap, Css.BOX_SHADOW_PERSPECTIVE_BOTTOM_LEFT_MARGIN_LEFT, Css.BOX_SHADOW_PERSPECTIVE_BOTTOM_LEFT_MARGIN_TOP, Css.BOX_SHADOW_PERSPECTIVE_BOTTOM_LEFT_MARGIN_RIGHT, Css.BOX_SHADOW_PERSPECTIVE_BOTTOM_LEFT_MARGIN_BOTTOM);
                    }else{
                        flag = false;
                    }
                    break;
                } case 22:{
                    if(Css.PERSPECTIVE_BOTTOM_RIGHT.equals(boxShadow)){
                        cssMap.put(Css.BOX_SHADOW, Css.BOX_SHADOW_PERSPECTIVE_BOTTOM_RIGHT);
                        setMaxMargin(cssMap, Css.BOX_SHADOW_PERSPECTIVE_BOTTOM_RIGHT_MARGIN_LEFT, Css.BOX_SHADOW_PERSPECTIVE_BOTTOM_RIGHT_MARGIN_TOP, Css.BOX_SHADOW_PERSPECTIVE_BOTTOM_RIGHT_MARGIN_RIGHT, Css.BOX_SHADOW_PERSPECTIVE_BOTTOM_RIGHT_MARGIN_BOTTOM);
                    }else{
                        flag = false;
                    }
                    break;
                } default:{
                    flag = false;
                }
            }

            if(!flag){
                //移除有关信息
                cssMap.remove(Css.BOX_SHADOW);
            }
        }

    }

    /**
     * 在css的margin四个方向属性中寻找到一个最大的赋值给他，在此操作之前应该已经做过了margin解析工作并且确保每一个方向的margin都是有值的
     * @param cssMap css原来的map
     * @param left 当前需要设定的左边界值
     * @param top 当前需要设定的上边界值
     * @param right 当前需要设定的右边界值
     * @param bottom 当前需要设定的下边界值
     */
    private static void setMaxMargin(Map<String, String> cssMap, int left, int top, int right, int bottom){
        if((int)Double.parseDouble(cssMap.get(Css.MARGIN_LEFT)) < left){
            cssMap.put(Css.MARGIN_LEFT, left + "");
        }
        if((int)Double.parseDouble(cssMap.get(Css.MARGIN_RIGHT)) < right){
            cssMap.put(Css.MARGIN_RIGHT, right + "");
        }
        if((int)Double.parseDouble(cssMap.get(Css.MARGIN_TOP)) < top){
            cssMap.put(Css.MARGIN_TOP, top + "");
        }
        if((int)Double.parseDouble(cssMap.get(Css.MARGIN_BOTTOM)) < bottom){
            cssMap.put(Css.MARGIN_BOTTOM, bottom + "");
        }
    }



    static void reflectionCssClean(Map<String, String> cssMap) {
        if(cssMap.containsKey(Css.REFLECTION)) {
            String reflection = cssMap.get(Css.REFLECTION).replace("-", "").replace("_", "").replace(" ", "");
            if(Css.NEAR_MIN.equalsIgnoreCase(reflection)){
                cssMap.put(Css.REFLECTION, Css.REFLECTION_NEAR_MIN);
                setMaxMargin(cssMap, Css.REFLECTION_NEAR_MIN_LEFT, Css.REFLECTION_NEAR_MIN_TOP, Css.REFLECTION_NEAR_MIN_RIGHT, Css.REFLECTION_NEAR_MIN_BOTTOM);
            }else if(Css.NEAR.equalsIgnoreCase(reflection)){
                cssMap.put(Css.REFLECTION, Css.REFLECTION_NEAR);
                setMaxMargin(cssMap, Css.REFLECTION_NEAR_LEFT, Css.REFLECTION_NEAR_TOP, Css.REFLECTION_NEAR_RIGHT, Css.REFLECTION_NEAR_BOTTOM);
            }else if(Css.NEAR_MAX.equalsIgnoreCase(reflection)){
                cssMap.put(Css.REFLECTION, Css.REFLECTION_NEAR_MAX);
                setMaxMargin(cssMap, Css.REFLECTION_NEAR_MAX_LEFT, Css.REFLECTION_NEAR_MAX_TOP, Css.REFLECTION_NEAR_MAX_RIGHT, Css.REFLECTION_NEAR_MAX_BOTTOM);
            }else if(Css.MEDIUM_MIN.equalsIgnoreCase(reflection)){
                cssMap.put(Css.REFLECTION, Css.REFLECTION_MEDIUM_MIN);
                setMaxMargin(cssMap, Css.REFLECTION_MEDIUM_MIN_LEFT, Css.REFLECTION_MEDIUM_MIN_TOP, Css.REFLECTION_MEDIUM_MIN_RIGHT, Css.REFLECTION_MEDIUM_MIN_BOTTOM);
            }else if(Css.MEDIUM.equalsIgnoreCase(reflection)){
                cssMap.put(Css.REFLECTION, Css.REFLECTION_MEDIUM);
                setMaxMargin(cssMap, Css.REFLECTION_MEDIUM_LEFT, Css.REFLECTION_MEDIUM_TOP, Css.REFLECTION_MEDIUM_RIGHT, Css.REFLECTION_MEDIUM_BOTTOM);
            }else if(Css.MEDIUM_MAX.equalsIgnoreCase(reflection)){
                cssMap.put(Css.REFLECTION, Css.REFLECTION_MEDIUM_MAX);
                setMaxMargin(cssMap, Css.REFLECTION_MEDIUM_MAX_LEFT, Css.REFLECTION_MEDIUM_MAX_TOP, Css.REFLECTION_MEDIUM_MAX_RIGHT, Css.REFLECTION_MEDIUM_MAX_BOTTOM);
            }else if(Css.FAR_MIN.equalsIgnoreCase(reflection)){
                cssMap.put(Css.REFLECTION, Css.REFLECTION_FAR_MIN);
                setMaxMargin(cssMap, Css.REFLECTION_FAR_MIN_LEFT, Css.REFLECTION_FAR_MIN_TOP, Css.REFLECTION_FAR_MIN_RIGHT, Css.REFLECTION_FAR_MIN_BOTTOM);
            }else if(Css.FAR.equalsIgnoreCase(reflection)){
                cssMap.put(Css.REFLECTION, Css.REFLECTION_FAR);
                setMaxMargin(cssMap, Css.REFLECTION_FAR_LEFT, Css.REFLECTION_FAR_TOP, Css.REFLECTION_FAR_RIGHT, Css.REFLECTION_FAR_BOTTOM);
            }else if(Css.FAR_MAX.equalsIgnoreCase(reflection)){
                cssMap.put(Css.REFLECTION, Css.REFLECTION_FAR_MAX);
                setMaxMargin(cssMap, Css.REFLECTION_FAR_MAX_LEFT, Css.REFLECTION_FAR_MAX_TOP, Css.REFLECTION_FAR_MAX_RIGHT, Css.REFLECTION_FAR_MAX_BOTTOM);
            }
        }
    }

    static void borderCssClean(Map<String, String> cssMap) {
        String[] borders = borderToClean(cssMap, 12700, 1270000, 100, "");
        if(borders != null){
            cssMap.put(Css.BORDER, "<a:ln w=\"" + borders[0] + "\"><a:solidFill><a:srgbClr val=\"" + borders[1] + "\"/></a:solidFill>" + borders[2] + "</a:ln>");
            int margin = (int)(Integer.parseInt(borders[0])*1.5);
            setMaxMargin(cssMap, margin, margin, margin, margin);
        }
    }

    /**
     * 解析tableMap生成用户行高控制数据
     * @param tableMap 用户tableMap数据
     * @return 返回用户的rowHeight数据
     */
    static Map<Integer, Integer> tableGetRowHeight(Map<String, String> tableMap, int paperHeight, int paddingTop,
                                                   int paddingBottom, int paperHead, int paperFoot) {
        //寻找是否有row-height
        if(tableMap.containsKey(Css.ROW_HEIGHT)){
            return getTableSize(tableMap, Css.ROW_HEIGHT, paperHeight - paddingTop -
                    paddingBottom - paperHead - paperFoot - 25);
        }else{
            return new HashMap<Integer, Integer>();
        }
    }

    /**
     * 获取用户的列宽信息，如果用户没有设置列宽，默认为平分，如果用户设置了列宽，那就设置的按设置来，其他平分
     *
     * @param length 一共多少列
     * @param tableMap  用户设定的cssMap
     * @return 返回从0开始，一共列宽个数字的List集合
     */
    static List<Integer> tableGetColumnWidth(int length, Map<String, String> tableMap, int paperWidth, int paddingLeft, int paddingRight) {
        Map<Integer, Integer> columnMap = new HashMap<Integer, Integer>();
        int allSize = paperWidth - paddingLeft - paddingRight - 11;
        if(tableMap.containsKey(Css.WIDTH)){
            String width = tableMap.get(Css.WIDTH).trim();
            if(width.contains(Css.PERCENT)){
                //百分数width表示
                width = width.replace(Css.PERCENT, "").trim();
                if(isNumeric(width)) {
                    int widthPercent = Math.abs((int) Double.parseDouble(width));
                    if(widthPercent > 0 && widthPercent <= 100){
                        allSize = (int)(allSize * widthPercent * 0.01);
                    }
                }
            }else{
                width = width.replace(Css.PT, "").replace(Css.PX, "").trim();
                //数值表示形式
                if(isNumeric(width)) {
                    allSize = Math.abs((int) Double.parseDouble(width)) * Docx.DOCX_PROPORTION;
                }
            }
        }
        if(tableMap.containsKey(Css.COLUMN_WIDTH)) {
            columnMap = getTableSize(tableMap, Css.COLUMN_WIDTH, allSize);
        }
        int other = length - columnMap.size();
        int nowSize = 0;
        int otherSize = 0;
        if(other != 0){
            for (Map.Entry<Integer, Integer> col : columnMap.entrySet()) {
                nowSize += col.getValue();
            }
            otherSize = (allSize - nowSize) / other;
        }
        List<Integer> column = new ArrayList<Integer>();
        for (int i = 0; i < length; i++) {
            if(columnMap.containsKey(i + 1)){
                column.add(columnMap.get(i + 1));
            }else{
                column.add(otherSize);
            }
        }
        return column;
    }

    /**
     * 提取的一个根据用户css信息寻找
     * @param tableMap css样式表解析出来的map集合
     * @param key 需要寻找的key值
     * @param allSize 整个的大小，比如要找横向宽度，就把横着总共多长传过来
     * @return 返回一个长度map集合，包含行号或列号 以及 需要设定的长度
     */
    private static Map<Integer, Integer> getTableSize(Map<String, String> tableMap, String key, int allSize){
        String[] sizes = tableMap.get(key).split(",");
        Map<Integer, Integer> sizeMap = new HashMap<Integer, Integer>();
        for (String size : sizes) {
            String[] lines = size.trim().split("=");
            if (lines.length == 2 && isNumeric(lines[0])) {
                int sizeId = Math.abs((int) Double.parseDouble(lines[0]));
                lines[1] = lines[1].trim();
                //有没有 %
                if (lines[1].contains(Css.PERCENT)) {
                    //百分数表示
                    lines[1] = lines[1].replace(Css.PERCENT, "").trim();
                    if (isNumeric(lines[1])) {
                        int sizeHeight = (int) (Math.abs((int) Double.parseDouble(lines[1])) * 0.01 * allSize);
                        sizeMap.put(sizeId, sizeHeight);
                    }
                } else {
                    //正常pt表示
                    lines[1] = lines[1].replace(Css.PX, "").replace(Css.PT, "").trim();
                    if (isNumeric(lines[1])) {
                        int sizeHeight = Math.abs((int) Double.parseDouble(lines[1])) * Docx.DOCX_PROPORTION;
                        sizeMap.put(sizeId, sizeHeight);
                    }
                }
            }
        }
        return sizeMap;
    }

    /**
     * 根据用户的设置数据设置表格线条样式
     * @param tableMap css的map集合
     * @return 返回border设置语句
     */
    static String tableGetBorder(Map<String, String> tableMap) {
        String[][] border = new String[6][3];
        //注意应该从小到大填充，先检查单独的
        if(tableMap.containsKey(Css.BORDER_TOP)){
            tableBorderGet(tableMap.get(Css.BORDER_TOP), border[0]);
        }
        if(tableMap.containsKey(Css.BORDER_LEFT)){
            tableBorderGet(tableMap.get(Css.BORDER_LEFT), border[1]);
        }
        if(tableMap.containsKey(Css.BORDER_BOTTOM)){
            tableBorderGet(tableMap.get(Css.BORDER_BOTTOM), border[2]);
        }
        if(tableMap.containsKey(Css.BORDER_RIGHT)){
            tableBorderGet(tableMap.get(Css.BORDER_RIGHT), border[3]);
        }
        if(tableMap.containsKey(Css.BORDER_INSIDE_H)){
            tableBorderGet(tableMap.get(Css.BORDER_INSIDE_H), border[4]);
        }
        if(tableMap.containsKey(Css.BORDER_INSIDE_V)){
            tableBorderGet(tableMap.get(Css.BORDER_INSIDE_V), border[5]);
        }
        //检查合并的
        if(tableMap.containsKey(Css.BORDER_INSIDE)){
            tableBorderGet(tableMap.get(Css.BORDER_INSIDE), border[4]);
            tableBorderGet(tableMap.get(Css.BORDER_INSIDE), border[5]);
        }
        if(tableMap.containsKey(Css.BORDER)){
            tableBorderGet(tableMap.get(Css.BORDER), border[0]);
            tableBorderGet(tableMap.get(Css.BORDER), border[1]);
            tableBorderGet(tableMap.get(Css.BORDER), border[2]);
            tableBorderGet(tableMap.get(Css.BORDER), border[3]);
            tableBorderGet(tableMap.get(Css.BORDER), border[4]);
            tableBorderGet(tableMap.get(Css.BORDER), border[5]);
        }
        //拼接字符串
        StringBuilder xml = tableBorderXmlDo(border);
        return xml.toString();
    }

    /**
     * 处理这一大堆数据，拼接成xml
     * @param border 拼合过的样式表内容数组
     * @return 返回StringBuilder类型的xml
     */
    private static StringBuilder tableBorderXmlDo(String[][] border) {
        StringBuilder xml = new StringBuilder();
        xml.append("<w:tblBorders><w:top w:val=\"");
        xml.append(tableBorderXmlDoText(border[0][0], "single"));
        xml.append("\" w:sz=\"");
        xml.append(tableBorderXmlDoText(border[0][1], "4"));
        xml.append("\" w:space=\"0\" w:color=\"");
        xml.append(tableBorderXmlDoText(border[0][2], "000000"));
        xml.append("\"/><w:left w:val=\"");
        xml.append(tableBorderXmlDoText(border[1][0], "single"));
        xml.append("\" w:sz=\"");
        xml.append(tableBorderXmlDoText(border[1][1], "4"));
        xml.append("\" w:space=\"0\" w:color=\"");
        xml.append(tableBorderXmlDoText(border[1][2], "000000"));
        xml.append("\"/><w:bottom w:val=\"");
        xml.append(tableBorderXmlDoText(border[2][0], "single"));
        xml.append("\" w:sz=\"");
        xml.append(tableBorderXmlDoText(border[2][1], "4"));
        xml.append("\" w:space=\"0\" w:color=\"");
        xml.append(tableBorderXmlDoText(border[2][2], "000000"));
        xml.append("\"/><w:right w:val=\"");
        xml.append(tableBorderXmlDoText(border[3][0], "single"));
        xml.append("\" w:sz=\"");
        xml.append(tableBorderXmlDoText(border[3][1], "4"));
        xml.append("\" w:space=\"0\" w:color=\"");
        xml.append(tableBorderXmlDoText(border[3][2], "000000"));
        xml.append("\"/><w:insideH w:val=\"");
        xml.append(tableBorderXmlDoText(border[4][0], "single"));
        xml.append("\" w:sz=\"");
        xml.append(tableBorderXmlDoText(border[4][1], "4"));
        xml.append("\" w:space=\"0\" w:color=\"");
        xml.append(tableBorderXmlDoText(border[4][2], "000000"));
        xml.append("\"/><w:insideV w:val=\"");
        xml.append(tableBorderXmlDoText(border[5][0], "single"));
        xml.append("\" w:sz=\"");
        xml.append(tableBorderXmlDoText(border[5][1], "4"));
        xml.append("\" w:space=\"0\" w:color=\"");
        xml.append(tableBorderXmlDoText(border[5][2], "000000"));
        xml.append("\"/></w:tblBorders>");
        return xml;
    }

    /**
     * 两个值返回一个有效的
     * @param border 期望值
     * @param normal 默认值
     * @return 有效值
     */
    private static String tableBorderXmlDoText(String border, String normal){
        if(border != null && !"".equals(border)){
            return border;
        }else{
            return normal;
        }
    }

    /**
     * 根据三段式、两段式、一段式自动识别边框标记的内容，颜色、线类型、数字可以任意位置书写，空格间隔就能自动识别
     * @param css css样式表的value
     */
    private static void tableBorderGet(String css, String[] border){
        String[] borders = css.trim().split(" ");
        int size = -1;
        for (String s : borders) {
            if(isNumeric(s.replace(Css.PX, "").replace(Css.PT, ""))){
                size = Math.abs((int)(Double.parseDouble(s) * 8));
                if(size >= 4 && size <= 64){
                    //接收0.5-8磅
                    if(border[1] == null || "".equals(border[1])) {
                        border[1] = size + "";
                    }
                }
            }
            String color = isColor(s);
            if(color != null){
                if(border[2] == null || "".equals(border[2])) {
                    border[2] = color;
                }
            }
            //也就是说上面都解析失败了
            boolean sizeLaw = size < 4 || size > 64;
            if(color == null && sizeLaw){
                for (int i = 0; i < Css.BORDER_TABLE_LOW.length; i++) {
                    s = s.replace("-", "").replace("_", "").trim().toLowerCase();
                    if(s.length() == Css.BORDER_TABLE_LOW[i].length() && Css.BORDER_TABLE_LOW[i].equals(s)){
                        if(border[0] == null || "".equals(border[0])) {
                            border[0] = Css.BORDER_TABLE[i];
                            break;
                        }
                    }
                }
            }
        }
    }

    /**
     * 根据用户的css数据确定对齐样式
     * @param tableMap 用户的css样式表
     * @return 返回用户的对齐信息代码
     */
    String tableGetTextAlign(Map<String, String> tableMap) {
        if(tableMap.containsKey(Css.TEXT_ALIGN)){
            if(Css.TEXT_ALIGN_CENTER.equals(tableMap.get(Css.TEXT_ALIGN))){
                return new IoUtil().readFileJar(File.separator + "align" + File.separator + "center.tbl").toString();
            }else if(Css.TEXT_ALIGN_RIGHT.equals(tableMap.get(Css.TEXT_ALIGN))){
                return new IoUtil().readFileJar(File.separator + "align" + File.separator + "right.tbl").toString();
            }
        }
        return "";
    }

    /**
     * 根据用户的样式选择添加样式到文件同时生成id，这里我们网格表采用 a-b 模板标签，清单表采用 a0-b 标签，默认两个模板分别使用a3或a4
     * @param tableMap css解析后的tableMap数据
     * @param tempPath 临时文件储存目录
     * @param uri 当前运行目录
     * @return 返回用户模板标签转化后标签名称
     */
    String tableGetTemplate(Map<String, String> tableMap, String tempPath, String uri) {
        String template = "normal1";
        if(tableMap.containsKey(Css.TEMPLATE)){
            String templateText = tableMap.get(Css.TEMPLATE);
            if(templateText.contains(Css.NORMAL)){
                if(isNumeric(templateText.substring(6))){
                    template = "a" + (Integer.parseInt(templateText.substring(6)) + 2);
                }
            }else if(templateText.contains(Css.GIRD) || templateText.contains(Css.LIST)){
                if(isNumeric(templateText.substring(4, 5)) && isNumeric(templateText.substring(5, 6))){
                    int a = Integer.parseInt(templateText.substring(4, 5));
                    int b = Integer.parseInt(templateText.substring(5, 6));
                    if(a > 0 && a < 8 && b > 0 && b < 8){
                        if(templateText.contains(Css.GIRD)){
                            template = a + "-" + b;
                        }else if(templateText.contains(Css.LIST)){
                            template = a + "0-" + b;
                        }
                    }
                }
            }
        }
        if("normal1".equals(template)){
            template = "a3";
        }
        //添加索引
        new IoUtil().addTemplate(template, tempPath);
        return template;
    }

    /**
     * 根据列的数值生成列宽代码
     * @param columnList 列宽List
     */
    static String tableGetColumnWidth(List columnList){
        StringBuilder xml = new StringBuilder();
        for (Object o : columnList) {
            xml.append("<w:gridCol w:w=\"");
            xml.append(o);
            xml.append("\"/>");
        }
        return xml.toString();
    }

    static String tableRowCnfStyle(int i){
        StringBuilder xml = new StringBuilder();
        int firstRow = 0;
        if(i == 0){firstRow = 1;}
        int lastRow = 0;
        //if(i == length - 1){lastRow = 1;}
        int oddBandH = 0;
        if(i % 2 == 1){oddBandH = 1;}
        int evenBandH = 0;
        if(i != 0 && i % 2 == 0){evenBandH = 1;}
        xml.append("<w:cnfStyle w:val=\"");
        xml.append(firstRow);
        xml.append(lastRow);
        xml.append("0000");
        xml.append(oddBandH);
        xml.append(evenBandH);
        xml.append("0000\" w:firstRow=\"");
        xml.append(firstRow);
        xml.append("\" w:lastRow=\"");
        xml.append(lastRow);
        xml.append("\" w:firstColumn=\"0\" w:lastColumn=\"0\" w:oddVBand=\"0\" w:evenVBand=\"0\" w:oddHBand=\"");
        xml.append(oddBandH);
        xml.append("\" w:evenHBand=\"");
        xml.append(evenBandH);
        xml.append("\" w:firstRowFirstColumn=\"0\" w:firstRowLastColumn=\"0\" w:lastRowFirstColumn=\"0\" w:lastRowLastColumn=\"0\"/>");
        return xml.toString();
    }

    static String tableColumnCnfStyle(int i){
        if(i == 0){
            return "<w:cnfStyle w:val=\"001000000000\" w:firstRow=\"0\" w:lastRow=\"0\" w:firstColumn=\"1\" w:lastColumn=\"0\" w:oddVBand=\"0\" w:evenVBand=\"0\" w:oddHBand=\"0\" w:evenHBand=\"0\" w:firstRowFirstColumn=\"0\" w:firstRowLastColumn=\"0\" w:lastRowFirstColumn=\"0\" w:lastRowLastColumn=\"0\"/>";
        }else{
            return "<w:cnfStyle w:val=\"000000000000\" w:firstRow=\"0\" w:lastRow=\"0\" w:firstColumn=\"0\" w:lastColumn=\"0\" w:oddVBand=\"0\" w:evenVBand=\"0\" w:oddHBand=\"0\" w:evenHBand=\"0\" w:firstRowFirstColumn=\"0\" w:firstRowLastColumn=\"0\" w:lastRowFirstColumn=\"0\" w:lastRowLastColumn=\"0\"/>";
        }
    }

    /**
     * 生成页眉和页脚信息
     * @param headMap 页眉css样式表
     * @param footMap 页脚css样式表
     * @return 返回拼接好的xml
     */
    StringBuilder paperOutGetStyle(Map<String, String> headMap, Map<String, String> footMap) {
        StringBuilder xml = new StringBuilder();
        if(headMap == null){
            xml.append(new IoUtil().readFileJar(Docx.DOCX_OUT_HEAD_URI));
        }else{
            xml.append(new IoUtil().readFileJar(Docx.DOCX_OUT_HEAD_DIY_URI));
            paperOutMapDo(headMap, xml, "bottom");
        }

        if(footMap == null){
            xml.append(new IoUtil().readFileJar(Docx.DOCX_OUT_FOOT_URI));
        }else{
            xml.append(new IoUtil().readFileJar(Docx.DOCX_OUT_FOOT_DIY_URI));
            paperOutMapDo(headMap, xml, "top");
        }
        return xml;
    }

    /**
     * 整理页眉页脚的map集合
     * @param cssMap css的map集合
     * @param xml 希望读写的xml
     * @param position 指定边框方向
     */
    private static void paperOutMapDo(Map<String, String> cssMap, StringBuilder xml, String position){
        String[] borders = borderToClean(cssMap, 8, 56, 4, "single");
        if(borders != null){
            replace(xml, Css.A_BORDER_A, "<w:" + position + " w:val=\"" + borders[2] + "\" w:sz=\"" + borders[0] +
                    "\" w:space=\"1\" w:color=\"" + borders[1] + "\"/>");
        }else{
            replace(xml, Css.A_BORDER_A, "");
        }
        if(cssMap.containsKey(Css.FONT_SIZE)){
            replace(xml, Css.A_FONT_SIZE_A, fontSizeToInt(cssMap.get(Css.FONT_SIZE)) + "");
        }else{
            replace(xml, Css.A_FONT_SIZE_A, "18");
        }
        if(cssMap.containsKey(Css.TEXT_ALIGN)){
            if(Css.TEXT_ALIGN_CENTER.equals(cssMap.get(Css.TEXT_ALIGN))) {
                replace(xml, Css.A_TEXT_ALIGN_A, "<w:jc w:val=\"center\"/>");
            }else if(Css.TEXT_ALIGN_RIGHT.equals(cssMap.get(Css.TEXT_ALIGN))) {
                replace(xml, Css.A_TEXT_ALIGN_A, "<w:jc w:val=\"right\"/>");
            }else{
                replace(xml, Css.A_TEXT_ALIGN_A, "");
            }
        }else{
            replace(xml, Css.A_TEXT_ALIGN_A, "");
        }
    }

    /**
     * border方法整合
     * @param cssMap css样式表文件
     * @param borderTimes border四级储存放大倍数
     * @param borderMax border最大是多少，写放大之后的
     * @param borderDefault border默认是多少，写放大之后的
     * @param typeDefault 默认的边框样式
     * @return 返回数组表示边框类型尺寸和颜色
     */
    private static String[] borderToClean(Map<String, String> cssMap, int borderTimes, int borderMax, int borderDefault, String typeDefault){
        String borderColor = null;
        String borderSize = null;
        String borderType = null;
        String typeXml = null;
        if(cssMap.containsKey(Css.BORDER)) {
            String[] borders = cssMap.get(Css.BORDER).replace("  ", " ").split(" ");
            if(borders.length > 0){
                borderSize = borders[0];
            }
            if(borders.length > 1){
                borderType = borders[1];
            }
            if(borders.length > 2){
                borderColor = borders[2];
            }
            cssMap.remove(Css.BORDER);
        }
        if(cssMap.containsKey(Css.BORDER_WIDTH)) {
            borderSize = cssMap.get(Css.BORDER_WIDTH);
            cssMap.remove(Css.BORDER_WIDTH);
        }
        if(cssMap.containsKey(Css.BORDER_STYLE)) {
            borderType = cssMap.get(Css.BORDER_STYLE);
            cssMap.remove(Css.BORDER_STYLE);
        }
        if(cssMap.containsKey(Css.BORDER_COLOR)) {
            borderColor = cssMap.get(Css.BORDER_COLOR);
            cssMap.remove(Css.BORDER_COLOR);
        }
        if(borderSize != null || borderType != null || borderColor != null){
            if(borderSize != null){
                borderSize = borderSize.replace("px", "").replace("pt", "");
                if(isNumeric(borderSize)){
                    int borderSizeNum = Math.abs((int)(Double.parseDouble(borderSize) * borderTimes));
                    if(borderSizeNum < 0 || borderSizeNum > borderMax){
                        borderSize = borderSizeNum + "";
                    }else{
                        borderSize = borderDefault + "";
                    }
                }else{
                    borderSize = borderDefault + "";
                }
            }
            borderColor = isColor(borderColor);
            if(borderColor == null){
                borderColor = "000000";
            }
            for (int i = 0; i < Css.BORDER_TYPE.length; i++) {
                if(Css.BORDER_TYPE[i].equals(borderType)){
                    typeXml = Css.BORDER_TYPE_XML[i];
                }
            }
            if(typeXml == null){
                typeXml = typeDefault;
            }
            return new String[]{borderSize, borderColor, typeXml};
        }else{
            return null;
        }
    }

    static StringBuilder coreDo(WordCore wordCore) {
        if(wordCore == null){
            wordCore = new WordCore();
        }
        StringBuilder xml = new StringBuilder();
        xml.append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n" +
                "<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"><dc:title>");
        if(wordCore.getTitle() != null && !"".equals(wordCore.getTitle())){
            xml.append(wordCore.getTitle().trim());
        }
        xml.append("</dc:title><dc:subject>");
        if(wordCore.getSubject() != null && !"".equals(wordCore.getSubject())){
            xml.append(wordCore.getSubject().trim());
        }
        xml.append("</dc:subject><dc:creator>");
        if(wordCore.getCreator() != null && !"".equals(wordCore.getCreator())){
            xml.append(wordCore.getCreator().trim());
        }
        xml.append("</dc:creator><cp:keywords>");
        if(wordCore.getKeywords() != null && !"".equals(wordCore.getKeywords())){
            xml.append(wordCore.getKeywords().trim());
        }
        xml.append("</cp:keywords><dc:description>");
        if(wordCore.getDescription() != null && !"".equals(wordCore.getDescription())){
            xml.append(wordCore.getDescription().trim());
        }
        xml.append("</dc:description><cp:lastModifiedBy>");
        if(wordCore.getLastModifiedBy() != null && !"".equals(wordCore.getLastModifiedBy())){
            xml.append(wordCore.getLastModifiedBy().trim());
        }
        xml.append("</cp:lastModifiedBy><cp:revision>");
        xml.append(wordCore.getRevision());
        xml.append("</cp:revision><dcterms:created xsi:type=\"dcterms:W3CDTF\">");
        SimpleDateFormat sdf1 = new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat sdf2 = new SimpleDateFormat("HH:mm:ss");
        if(wordCore.getCreated() == null){
            wordCore.setCreated(new Date());
        }
        if(wordCore.getModified() == null){
            wordCore.setModified(new Date());
        }
        xml.append(sdf1.format(wordCore.getCreated()));
        xml.append("T");
        xml.append(sdf2.format(wordCore.getCreated()));
        xml.append("Z</dcterms:created><dcterms:modified xsi:type=\"dcterms:W3CDTF\">");
        xml.append(sdf1.format(wordCore.getModified()));
        xml.append("T");
        xml.append(sdf2.format(wordCore.getModified()));
        xml.append("Z</dcterms:modified></cp:coreProperties>");
        return xml;
    }
}
