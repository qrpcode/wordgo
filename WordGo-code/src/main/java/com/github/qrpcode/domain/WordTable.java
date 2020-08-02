package com.github.qrpcode.domain;

import com.github.qrpcode.config.Docx;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 这是一个word中表格的类，主要用于记录表格操作信息，只有被add到wordGo中才会开始生成表格代码
 * @author QiRuipeng
 */
public class WordTable {

    private StringBuilder[][] textArr;
    private StringBuilder[][] cssArr;
    private StringBuilder tableCss;
    private List<String> img;
    private List<String> imgCss;
    private List<String> text;
    private List<String> css;

    /**
     * 创建一个几行几列的表格
     * @param row 行数，输入数字即为实际行数，计数从1开始
     * @param column 列数，输入数字即为实际行数，计数从1开始
     */
    public WordTable(int row, int column){
        textArr = new StringBuilder[row][column];
        cssArr = new StringBuilder[row][column];
        img = new ArrayList<String>();
        css = new ArrayList<String>();
        text = new ArrayList<String>();
        imgCss = new ArrayList<String>();
        tableCss = new StringBuilder(" ");
    }

    /**
     * 创建一个几行几列的表格，并设置样式
     * 表格默认样式是宽度高度自动调整，横向宽度是100%
     * @param row  行数，输入数字即为实际行数，计数从1开始
     * @param column  列数，输入数字即为实际行数，计数从1开始
     * @param css  样式表
     */
    public WordTable(int row, int column, String css){
        this(row, column);
        tableCss = new StringBuilder(css);
    }

    public StringBuilder[][] getTextArr() {
        return textArr;
    }

    public StringBuilder[][] getCssArr() {
        return cssArr;
    }

    public StringBuilder getTableCss() {
        return tableCss;
    }

    public List<String> getImg() {
        return img;
    }

    public List<String> getImgCss() {
        return imgCss;
    }

    public List<String> getCss() {
        return css;
    }

    public List<String> getText() {
        return text;
    }

    /**
     * 将一些单元格合并起来，如果其中只有一个表格被填充了数据，将会
     * 合并失败可能的原因：单元格不存在、包含且未完全包含其他已经合并单元格
     * @param rowLeftTop 左上角单元格行号，从1开始数
     * @param columnLeftTop 左上角单元格列号，从1开始数
     * @param rowRightBottom 右下角单元格行号，从1开始数
     * @param columnRightBottom 右下角单元格列号，从1开始数
     * @return 如果合并失败返回false，合并成功返回true。
     */
    public boolean merge(int rowLeftTop, int columnLeftTop, int rowRightBottom, int columnRightBottom){
        //判断合理性
        if(rowLeftTop < 1){ rowLeftTop = 1; }
        if(columnLeftTop < 1){ rowLeftTop = 1; }
        if(rowRightBottom > textArr.length){ rowRightBottom = textArr.length;}
        if(columnRightBottom > textArr[0].length){ rowRightBottom = textArr[0].length;}
        if(rowRightBottom - rowLeftTop < 0 || columnRightBottom - columnLeftTop < 0 ||
                (rowLeftTop == rowRightBottom && columnLeftTop == columnRightBottom)){
            return false;
        }
        //因为我们是从1开始，实际上储存是从0开始，这里统一减1
        rowLeftTop--;
        rowRightBottom--;
        columnLeftTop--;
        columnRightBottom--;
        //用一个Map集合储存确保相同合并值我们只遍历一次，要不太慢了
        Map<String, Boolean> map = new HashMap<String, Boolean>();
        //第一次循环检查是否合并的合理
        for (int i = rowLeftTop; i <= rowRightBottom; i++) {
            for (int j = columnLeftTop; j <= columnRightBottom; j++) {
                if(textArr[i][j] != null && !map.containsKey(textArr[i][j].toString()) && textArr[i][j].indexOf(Docx.DOCX_TABLE_KEY) > -1){
                    map.put(textArr[i][j].toString(), true);
                    //这个单元格是一个合并单元格
                    String[] tables = textArr[i][j].toString().split("#");
                    int rowLeftTop2 = 0;
                    int columnLeftTop2 = 0;
                    int rowRightBottom2 = 0;
                    int columnRightBottom2 = 0;
                    if(Integer.parseInt(tables[1]) > i || Integer.parseInt(tables[2]) > j){
                        //说明这是单元格的左上角格子
                        rowLeftTop2 = i;
                        columnLeftTop2 = j;
                        rowRightBottom2 = Integer.parseInt(tables[1]);
                        columnRightBottom2 = Integer.parseInt(tables[2]);
                    }else{
                        //说明这是合并单元格中的一项
                        rowLeftTop2 = Integer.parseInt(tables[1]);
                        columnLeftTop2 = Integer.parseInt(tables[2]);
                        String[] tables2 = textArr[rowLeftTop2][columnLeftTop2].toString().split("#");
                        rowRightBottom2 = Integer.parseInt(tables2[1]);
                        columnRightBottom2 = Integer.parseInt(tables2[2]);
                    }
                    //如果左上右下任意一个超过这个边界我们就返回false
                    if(rowLeftTop2 < rowLeftTop || rowRightBottom2 > rowRightBottom || columnLeftTop2 < columnLeftTop ||
                            columnRightBottom2 > columnRightBottom){
                        return false;
                    }
                }
            }
        }
        //处理左上角格子数据
        if (textArr[rowLeftTop][columnLeftTop] != null && textArr[rowLeftTop][columnLeftTop].indexOf(Docx.DOCX_TABLE_KEY) > -1){
            String[] tables = textArr[rowLeftTop][columnLeftTop].toString().split("#");
            textArr[rowLeftTop][columnLeftTop] = new StringBuilder(tables[3]);
        }
        if (textArr[rowLeftTop][columnLeftTop] == null){
            textArr[rowLeftTop][columnLeftTop] = new StringBuilder(Docx.DOCX_TABLE_KEY + rowRightBottom +
                    Docx.DOCX_TABLE_HASH + columnRightBottom + Docx.DOCX_TABLE_HASH);
        }else{
            textArr[rowLeftTop][columnLeftTop] = textArr[rowLeftTop][columnLeftTop].insert(0, Docx.DOCX_TABLE_KEY +
                    rowRightBottom + Docx.DOCX_TABLE_HASH + columnRightBottom + Docx.DOCX_TABLE_HASH);
        }
        //第二次循环填充数据
        for (int i = rowLeftTop; i <= rowRightBottom; i++) {
            for (int j = columnLeftTop; j <= columnRightBottom; j++) {
                if(!(i == rowLeftTop && j == columnLeftTop)) {
                    //如果单元格是已合并单元格的左上角格子，是可能有数据的
                    if (textArr[i][j] != null && textArr[i][j].indexOf(Docx.DOCX_TABLE_KEY) > -1) {
                        String[] tables = textArr[i][j].toString().split("#");
                        if (tables.length > 3 && tables[3] != null && !"".equals(tables[3]) &&
                                (Integer.parseInt(tables[1]) > i || Integer.parseInt(tables[2]) > j)) {
                            textArr[i][j] = new StringBuilder(tables[3]);
                        }
                    }else if (textArr[i][j] != null && !"".equals(textArr[i][j].toString())){
                        //正常格子，但不能是左上角的格子
                        textArr[rowLeftTop][columnLeftTop].append(Docx.DOCX_TABLE_CUT);
                        textArr[rowLeftTop][columnLeftTop].append(textArr[i][j].toString());
                        cssArr[rowLeftTop][columnLeftTop].append(Docx.DOCX_TABLE_CUT);
                        if (cssArr[i][j] != null && !"".equals(cssArr[i][j].toString())){
                            cssArr[rowLeftTop][columnLeftTop].append(cssArr[i][j].toString());
                        }else{
                            cssArr[rowLeftTop][columnLeftTop].append(" ");
                        }
                    }
                    //将格子置为占位附属单元格
                    textArr[i][j] = new StringBuilder(Docx.DOCX_TABLE_KEY + rowLeftTop + Docx.DOCX_TABLE_HASH + columnLeftTop);
                    cssArr[i][j] = null;
                }
            }
        }
        //（PS 这种写法似乎不大好，但是一时间想不到其他写法了）
        return true;
    }

    /**
     * 从第几行第几列开始向左侧填充数据，如果其中包含合并单元格将会只填充一次。
     * 比如开始这一行有四个格子，第二个和第三个合并了，add(1, 1, "a", "b", "c");
     * 将会在第一个格子填充 a，第二个和第三个填充 b，第四个格子填充c
     *
     * @param row  行号，输入数字即为生活中的行号，也就是计数从1开始数
     * @param column  列号，输入数字即为生活中的列号，也就是计数从1开始数
     * @param str 需要填充数据的文本
     */
    public void add(int row, int column, String ... str){
        row = row - 1;
        column = column - 1;
        if(str != null){
            for (int i = 0; i < str.length; i++) {
                //先效验确定表格坐标合法
                if(row < textArr.length && (column + i) < textArr[0].length){
                    String[] tables = null;
                    while(textArr[row][column + i] != null && textArr[row][column + i].indexOf(Docx.DOCX_TABLE_KEY) > -1){
                        tables = textArr[row][column + i].toString().split("#");
                        if((Integer.parseInt(tables[1]) > row || Integer.parseInt(tables[2]) > column + i)) {
                            break;
                        }else{
                            column++;
                        }
                    }

                    if(textArr[row][column + i] == null){
                        textArr[row][column + i] = new StringBuilder(str[i].replace("#", ""));
                    }else{
                        textArr[row][column + i].append(Docx.DOCX_TABLE_HASH);
                        textArr[row][column + i].append(str[i].replace("#", ""));
                    }
                }
            }
        }
    }

    /**
     * 向表格中插入一个图片
     * @param row  行号，输入数字即为生活中的行号，也就是计数从1开始数
     * @param column  列号，输入数字即为生活中的列号，也就是计数从1开始数
     * @param uri 图片地址，必须是本地图片，支持jpg png gif
     * @param css 图片样式
     */
    public boolean addImg(int row, int column, String uri, String css){
        return addContent(row, column, uri, css, img, imgCss, Docx.DOCX_TABLE_IMG_KEY);
    }

    /**
     * 添加一段具有独特样式的文字，和该方格css定义重合部分会以在这里定义的独特样式为准，其他没重新定义的继承此窗格的样式
     * @param row  行号，输入数字即为生活中的行号，也就是计数从1开始数
     * @param column  列号，输入数字即为生活中的列号，也就是计数从1开始数
     * @param text 文本内容
     * @param css 文本样式
     */
    public boolean addStyleText(int row, int column, String text, String css){
        return addContent(row, column, text.replace("#", ""), css, this.text, this.css, Docx.DOCX_TABLE_CSS_KEY);
    }

    /**
     * 添加图片和css存在大量相同代码，现在进行一下抽取
     * @param row  行号，输入数字即为生活中的行号，也就是计数从1开始数
     * @param column 列号，输入数字即为生活中的列号，也就是计数从1开始数
     * @param content 内容，图片URI或者文字
     * @param css 文本样式
     * @param contentList 储存内容的list
     * @param cssList 储存样式的list
     * @param key 插入内容前缀
     * @return 插入成功了没
     */
    private boolean addContent(int row, int column, String content, String css, List<String> contentList, List<String> cssList, String key){
        row = row - 1;
        column = column - 1;
        if(row < textArr.length && column < textArr[0].length){
            if(textArr[row][column] == null){
                contentList.add(content);
                cssList.add(css);
                textArr[row][column] = new StringBuilder(key + (contentList.size() - 1));
                return true;
            }else{
                if(textArr[row][column].indexOf(Docx.DOCX_TABLE_KEY) > -1){
                    String[] tables = textArr[row][column].toString().split("#");
                    if(!(Integer.parseInt(tables[1]) > row || Integer.parseInt(tables[2]) > column)) {
                        return false;
                    }
                }
                contentList.add(content);
                cssList.add(css);
                textArr[row][column].append(Docx.DOCX_TABLE_HASH);
                textArr[row][column].append(key);
                textArr[row][column].append(contentList.size() - 1);
                return true;
            }
        }else{
            return false;
        }
    }

    /**
     * 设置表格样式
     * @param row  行号，输入数字即为生活中的行号，也就是计数从1开始数
     * @param column  列号，输入数字即为生活中的列号，也就是计数从1开始数
     * @param css 样式
     */
    public void addStyle(int row, int column, String css){
        row = row - 1;
        column = column - 1;
        if(row < textArr.length && column < textArr[0].length){
            if(cssArr[row][column] == null){
                cssArr[row][column] = new StringBuilder(css);
            }else{
                cssArr[row][column].append(";");
                cssArr[row][column].append(css);
            }
        }
    }

    /**
     * 批量设置表格样式
     * @param rowLeftTop 左上角单元格行号，从1开始数
     * @param columnLeftTop 左上角单元格列号，从1开始数
     * @param rowRightBottom 右下角单元格行号，从1开始数
     * @param columnRightBottom 右下角单元格列号，从1开始数
     * @param css 样式
     */
    public void addStyle(int rowLeftTop, int columnLeftTop, int rowRightBottom, int columnRightBottom, String css){
        //判断合理性
        if(rowLeftTop < 1){ rowLeftTop = 1; }
        if(columnLeftTop < 1){ rowLeftTop = 1; }
        if(rowRightBottom > textArr.length){ rowRightBottom = textArr.length;}
        if(columnRightBottom > textArr[0].length){ rowRightBottom = textArr[0].length;}
        //转化0开始储存模式
        rowLeftTop--;
        columnLeftTop--;
        rowRightBottom--;
        columnRightBottom--;
        //开始遍历储存
        for (int i = rowLeftTop; i <= rowRightBottom; i++) {
            for (int j = columnLeftTop; j <= columnRightBottom; j++) {
                if(cssArr[i][j] == null){
                    cssArr[i][j] = new StringBuilder(css);
                }else{
                    cssArr[i][j].append(";");
                    cssArr[i][j].append(css);
                }
            }
        }
    }
}
