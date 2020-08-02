package com.github.qrpcode.util;

import com.github.qrpcode.config.Css;
import com.github.qrpcode.config.Docx;
import com.github.qrpcode.domain.WordCore;
import com.github.qrpcode.domain.WordTable;

import java.io.File;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutorService;

/**
 * 初始生成一个word文件
 * 为了还在用1.5的人们...我们不能写@override
 * @author QiRuipeng
 */
public class WordFileUtilImpl implements WordFileUtil{

    /**
     * 初始化一个Word模板
     * @return 返回临时目录
     */
    public String startWordGo(){
        return new IoUtil().buildUtil(new IoUtil().getUri());
    }

    /**
     * 封装文本数据
     * @param text 需要封装的文本
     * @param css 需要的样式表
     * @return 返回封装好的 w:r 标签文本
     */
    public StringBuilder wordTestDo(String text, String css, ExecutorService es){
        Map<String, String> map = StrUtil.cssToMap(css);
        StringBuilder xml;

        //因为 text-decoration-color 属性需要嵌套 text-decoration 标签，提前解析。
        //如果设定颜色却没设定下划线，直接忽略这个属性
        if(map.containsKey(Css.TEXT_DECORATION_COLOR) && map.containsKey(Css.TEXT_DECORATION)){
            xml = TextUtil.decorationColor(new StringBuilder(Docx.DOC_TEXT), map.get(Css.TEXT_DECORATION_COLOR));
            map.remove(Css.TEXT_DECORATION_COLOR);
        }else if(map.containsKey(Css.TEXT_DECORATION)){
            xml = TextUtil.textDecoration(new StringBuilder(Docx.DOC_TEXT));
        }else{
            xml = new StringBuilder(Docx.DOC_TEXT);
        }

        //提前记录size防止多线程删除改变长度
        int mapSize = map.size();
        TextUtil textUtil = new TextUtil(xml, map);
        for (int i = 0; i < mapSize; i++) {
            es.submit(textUtil);
        }

        //多线程执行的是否每解析一个就会删除一个，我们需要等待都执行完在进行其他操作
        while(textUtil.getMap().size() != 0){
            try {
                Thread.sleep(1);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }

        StringBuilder cssXml = textUtil.getXml();
        //清除未匹配内容
        StrUtil.cleanNoMarry(cssXml);
        //文本必须最后导入防止影响替换
        StrUtil.replace(cssXml, "#text#", StrUtil.textChoose(text).toString());
        return cssXml;
    }

    /**
     * 切换到新的一页
     * @return 返回xml
     */
    public StringBuilder wordNewPageDo(){
        return new StringBuilder(Docx.DOCX_NEW_LINE);
    }

    /**
     * 拼接合成一行的文本
     * @param lineXml 这一行中已经生成的xml
     * @param css 这一行中所有样式表
     * @return 返回一行合成好的xml
     */
    public StringBuilder wordLineDo(StringBuilder lineXml, StringBuilder css) {
        //这里只需要解析两个标签，一个是 line-height 另一个是 text-align
        //第一步：切割css寻找行高、对齐方式、最大字号
        String cssStr = css.toString();
        String [] cssArray = cssStr.split(";");
        //必须倒序遍历，才能保留最后的设置
        int fontSize = -1;
        String lineHeight = null;
        String textAlign = Css.TEXT_ALIGN_LEFT;
        for (int i = cssArray.length - 1 ; i >= 0 ; i--) {
            if(cssArray[i].toLowerCase().contains(Css.LINE_HEIGHT)){
                String[] lineKeyValue = cssArray[i].toLowerCase().split(":");
                lineHeight = lineKeyValue[1].trim();
            }else if(cssArray[i].toLowerCase().contains(Css.TEXT_ALIGN)){
                String[] alignKeyValue = cssArray[i].toLowerCase().split(":");
                if(Css.TEXT_ALIGN_RIGHT.equals(alignKeyValue[1].trim())){
                    textAlign = Css.TEXT_ALIGN_RIGHT;
                }else if(Css.TEXT_ALIGN_CENTER.equals(alignKeyValue[1].trim())){
                    textAlign = Css.TEXT_ALIGN_CENTER;
                }
            }else if(cssArray[i].toLowerCase().contains(Css.FONT_SIZE)){
                String[] sizeKeyValue = cssArray[i].toLowerCase().split(":");
                if(StrUtil.fontSizeToInt(sizeKeyValue[1]) > fontSize){
                    fontSize = StrUtil.fontSizeToInt(sizeKeyValue[1]);
                }
            }
        }
        //如果用户没设置字体但是设置了行距，将字体锁定为五号字
        if(lineHeight != null && fontSize == -1){
            fontSize = 21;
        }
        //拼装参数
        StringBuilder xml = new StringBuilder(Docx.DOC_LINE);
        if(lineHeight != null){
            StrUtil.replace(xml, Css.A_LINE_HEIGHT_A, TextUtil.lineHeight(lineHeight, fontSize));
        }
        if(!Css.TEXT_ALIGN_LEFT.equals(textAlign)){
            StrUtil.replace(xml, Css.A_TEXT_ALIGN_A, "<w:jc w:val=\"" + textAlign + "\"/>");
        }
        //清除未匹配内容
        StrUtil.cleanNoMarry(xml);
        StrUtil.replace(xml, "#lineXml#", lineXml.toString());
        return xml;
    }

    public Map<String, String> imgCssDo(String css){
        return StrUtil.imgCssToMap(css);
    }

    /**
     * 对图片进行数据处理
     * @param imgUri 图片uri
     * @param cssMap css切好的map集合
     * @param tempPath 临时目录
     * @param es 线程池
     * @return 返回处理好的xml文件
     */
    public StringBuilder wordImgDo(String imgUri, Map<String, String> cssMap, int paperWidth, int paddingLeft, int paddingRight,
                                   int paperHeight, int paperHead, int paperBottom, String tempPath, ExecutorService es){
        StringBuilder xml;
        //第一步：拷贝文件
        int imgId = IoUtil.newImgAdd(tempPath, imgUri);
        //第二步：添加配置
        IoUtil.addContentTypes(tempPath, imgUri);
        //第三步：添加文件资源目录获取资源号
        int relNum = IoUtil.addRelationship(tempPath, imgUri, imgId);
        //第四步：规范化 css 属性，补全必要参数
        StrUtil.positionCssClean(cssMap, paperWidth, paddingLeft, paddingRight, paperHeight, paperHead, paperBottom);
        StrUtil.sizeCssClean(cssMap, imgUri);
        StrUtil.boxShadowCssClean(cssMap);
        StrUtil.reflectionCssClean(cssMap);
        StrUtil.borderCssClean(cssMap);
        //第五步：判断写入哪一个预设xml
        if(Css.POSITION_STATIC.equals(cssMap.get(Css.POSITION))){
            xml = new StringBuilder(Docx.DOC_IMG_STATIC);
        }else if(Css.POSITION_ABSOLUTE.equals(cssMap.get(Css.POSITION))){
            xml = new StringBuilder(Docx.DOC_IMG_ABSOLUTE);
        }else{
            xml = new StringBuilder(Docx.DOC_IMG_FIXED);
        }
        //第六步：开始写入
        int mapSize = cssMap.size();
        ImgUtil imgUtil = new ImgUtil(xml, cssMap);
        for (int i = 0; i < mapSize; i++) {
            es.submit(imgUtil);
        }
        //第七步：等待完成
        while(imgUtil.getMap().size() != 0){
            try {
                Thread.sleep(1);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
        }
        //第八步：清理没有赋值的标签
        StrUtil.cleanImgNoMarry(xml);
        //第九步：替换图片资源标签
        StrUtil.replace(xml, "#imgId#", imgId + "");
        StrUtil.replace(xml, "#rid#", relNum + "");
        return xml;
    }

    /**
     * 根据表格对象生成对应的xml代码
     * @param wordTable 表格table对象
     * @param tempPath 临时路径
     * @param es 线程池
     * @return 返回生成好的xml
     */
    public StringBuilder wordTableDo(WordTable wordTable, String tempPath, ExecutorService es,
                                     int paperHeight, int paddingTop, int paddingBottom, int paperHead, int paperFoot,
                                     int paperWidth, int paddingLeft, int paddingRight){
        String uri = new IoUtil().getUri();
        //第一步：初始化xml和map
        StringBuilder xml = new StringBuilder(Docx.DOCX_TABLE_HEAD);
        Map<String, String> tableMap = StrUtil.cssToMap(wordTable.getTableCss().toString());
        //第二步：解析表格宽高信息，行高默认自动无需赋值，列宽必须有值
        Map<Integer, Integer> rowMap = StrUtil.tableGetRowHeight(tableMap, paperHeight, paddingTop, paddingBottom, paperHead, paperFoot);
        List<Integer> columnList = StrUtil.tableGetColumnWidth(wordTable.getTextArr()[0].length, tableMap, paperWidth, paddingLeft, paddingRight);
        //第二步：解析tableMap参数，返回拼装好的信息
        StrUtil strUtil = new StrUtil();
        String template = strUtil.tableGetTemplate(tableMap, tempPath, uri);
        StrUtil.replace(xml, "@template@", template);
        StrUtil.replace(xml, "@text-align@", strUtil.tableGetTextAlign(tableMap));
        if(Css.BORDER_A3.equals(template)){
            StrUtil.replace(xml, "@tblBorders@", StrUtil.tableGetBorder(tableMap));
        }else{
            StrUtil.replace(xml, "@tblBorders@", "");
        }
        StrUtil.replace(xml, "@column-width@", StrUtil.tableGetColumnWidth(columnList));
        //第三步：清理没有赋值的标签，这里所有标签都已经强制赋值
        //第四步：双层循环生成表格同时生成文本内容
        StringBuilder[][] textArr = wordTable.getTextArr();
        StringBuilder[][] cssArr = wordTable.getCssArr();
        for (int i = 0; i < textArr.length; i++) {
            //4.1 添加行开始标签
            xml.append(Docx.DOCX_TABLE_ROW_1);
            //4.2 添加行高标签或者颜色标记标签
            xml.append(Docx.DOCX_TABLE_TRPR_1);
            //行高标签
            if(rowMap.containsKey(i + 1)) {
                xml.append(Docx.DOCX_TABLE_ROW_WIDTH_1);
                xml.append(rowMap.get(i + 1));
                xml.append(Docx.DOCX_TABLE_ROW_WIDTH_2);
            }
            //颜色标记
            xml.append(StrUtil.tableRowCnfStyle(i));
            xml.append(Docx.DOCX_TABLE_TRPR_2);
            //4.3 添加列标签
            for (int j = 0; j < textArr[0].length; j++) {
                //4.3.2 解析文本
                xml.append(tableGridDo(textArr[i][j], cssArr[i][j], wordTable.getImg(), wordTable.getImgCss(), wordTable.getText(),
                        wordTable.getCss(), columnList, i, j, textArr, tempPath, es, paperWidth, paddingLeft, paddingRight, paperHeight, paperHead, paddingBottom));
            }
            //4.4 添加结束标签
            xml.append(Docx.DOCX_TABLE_ROW_2);
        }
        //第五步：拼接尾部
        xml.append(Docx.DOCX_TABLE_FOOT);
        return xml;
    }

    /**
     * 根据表格的一个小格子进行生成代码
     * @param content 小格子文本，可以是跳过单元格标记信息
     * @param contentCss 小格子css
     * @param img 图片资源列表
     * @param imgCss 图片资源css样式表
     * @param text 样式文本列表
     * @param css css样式列表
     * @param columnList 表格宽度列表
     * @param textArr 表格数据
     * @param tempPath 临时路径
     * @param es 线程池
     * @return 返回处理好的xml
     */
    private static String tableGridDo(StringBuilder content, StringBuilder contentCss, List<String> img, List<String> imgCss,
                                      List<String> text, List<String> css, List<Integer> columnList, int row, int column,
                                      StringBuilder[][] textArr, String tempPath, ExecutorService es, int paperWidth,
                                      int paddingLeft, int paddingRight, int paperHeight, int paperHead, int paperBottom){
        StringBuilder xml = new StringBuilder();
        int columnWidth = 0;
        int gridSpan = 0;
        String vMerge = "";
        //判断是不是横向合并标签
        if(content != null && content.indexOf(Docx.DOCX_TABLE_KEY) > -1){
            String[] tables = content.toString().split("#");
            if(Integer.parseInt(tables[1]) > row || Integer.parseInt(tables[2]) > column) {
                //说明这是单元格的左上角格子
                if(Integer.parseInt(tables[1]) > row){
                    //说明存在跨行合并
                    vMerge = "<w:vMerge w:val=\"restart\"/>";
                }
                if(Integer.parseInt(tables[2]) > column){
                    for (int i = 0; i <= (Integer.parseInt(tables[2]) - column); i++) {
                        columnWidth += columnList.get(Integer.parseInt(tables[2]) + i);
                    }
                    gridSpan = Integer.parseInt(tables[2]) - column;
                }else{
                    columnWidth = columnList.get(column);
                }
                //添加tcp数据开始
                xml.append(Docx.DOCX_TABLE_COLUMN_WIDTH_1);
                xml.append(columnWidth);
                xml.append(Docx.DOCX_TABLE_COLUMN_WIDTH_2);
                if(gridSpan > 0){
                    xml.append("<w:gridSpan w:val=\"");
                    xml.append((gridSpan + 1));
                    xml.append("\"/>");
                }
                xml.append(vMerge);
                xml.append(StrUtil.tableColumnCnfStyle(column));
                xml.append("</w:tcPr>");
                //content截掉提示头
                if(tables.length > 3){
                    content = new StringBuilder();
                    for (int i = 3; i < tables.length; i++) {
                        content.append(tables[i]);
                        content.append("#");
                    }
                }else{
                    content = null;
                }
            }else{
                //说明是附属合并单元格
                if(Integer.parseInt(tables[1]) < row && Integer.parseInt(tables[2]) == column){
                    //这个是左上角单元格正下方单元格
                    int relRow = Integer.parseInt(tables[1]);
                    int relColumn = Integer.parseInt(tables[2]);
                    String[] goes = textArr[relRow][relColumn].toString().split("#");
                    if(Integer.parseInt(goes[2]) > relColumn){
                        for (int i = 0; i <= (Integer.parseInt(goes[2]) - relColumn); i++) {
                            columnWidth += columnList.get(Integer.parseInt(goes[2]) + i);
                        }
                        gridSpan = Integer.parseInt(goes[2]) - column;
                    }
                    xml.append("<w:tc><w:tcPr><w:tcW w:w=\"");
                    xml.append(columnWidth);
                    xml.append("\" w:type=\"dxa\"/><w:gridSpan w:val=\"");
                    xml.append(gridSpan + 1);
                    xml.append("\"/><w:vMerge/>");
                    xml.append(StrUtil.tableColumnCnfStyle(column));
                    xml.append("</w:tcPr><w:p w:rsidR=\"00AD336E\" w:rsidRDefault=\"00AD336E\"/>");
                    xml.append("</w:tc>");
                    return xml.toString();
                }else{
                    //这个是右侧单元格而且不在正下方
                    return "";
                }
            }
        }else{
            //这就是一个正常的文本单元格
            xml.append(Docx.DOCX_TABLE_COLUMN_WIDTH_1);
            xml.append(columnList.get(column));
            xml.append(Docx.DOCX_TABLE_COLUMN_WIDTH_2);
            xml.append(StrUtil.tableColumnCnfStyle(column));
            xml.append("</w:tcPr>");
        }
        //如果单元格为空直接返回
        if(content == null || "".equals(content.toString())){
            return xml.toString() + "<w:p w:rsidR=\"00AD336E\" w:rsidRDefault=\"00AD336E\"/></w:tc>";
        }else{
            xml.append(textToXml(content, contentCss, img, imgCss, text, css, tempPath, es, paperWidth, paddingLeft,
                    paddingRight, paperHeight, paperHead, paperBottom));
        }
        xml.append("</w:tc>");
        return xml.toString();
    }


    /**
     * 表格、页眉页脚解析文本
     * @param content 拼接字符串
     * @param contentCss 样式表
     * @param img 图片列表
     * @param imgCss 图片样式表列表
     * @param text 富文本表
     * @param css 富文本样式表列表
     * @param tempPath 临时路径
     * @param es 线程池
     * @return 合并xml
     */
    private static String textToXml(StringBuilder content, StringBuilder contentCss, List<String> img, List<String> imgCss,
                             List<String> text, List<String> css, String tempPath, ExecutorService es, int paperWidth,
                                    int paddingLeft, int paddingRight, int paperHeight, int paperHead, int paperBottom){
        StringBuilder xml = new StringBuilder();
        StringBuilder lineXml = new StringBuilder();
        String con = content.toString().replaceAll("(\\r\\n|\\n|\\n\\r)","&:&[br]&:&");
        con = con.replace("#", "&:&");
        String[] texts = con.split("&:&");
        for (String s : texts) {
            s = s.trim();
            if(s.length() > 0 && !"[br]".equals(s)){
                //行内信息处理
                if(s.length() > 4 && s.substring(0, 4).equals(Docx.DOCX_TABLE_CSS_KEY)){
                    int id = Integer.parseInt(s.substring(4));
                    lineXml.append(new WordFileUtilImpl().wordTestDo(text.get(id), css.get(id), es));
                }else if(s.length() > 4 && s.substring(0, 4).equals(Docx.DOCX_PAGE)){
                    //页码文本
                    int id = Integer.parseInt(s.substring(4));
                    String page = new WordFileUtilImpl().wordTestDo(Docx.DOCX_PAGE_TEXT, css.get(id), es).toString();
                    String[] pageDo = {};
                    if("template:1".equals(css.get(id))){
                        pageDo = Docx.DOCX_PAGE_1;
                    }else if("template:2".equals(css.get(id))){
                        pageDo = Docx.DOCX_PAGE_2;
                    }else if("template:3".equals(css.get(id))){
                        pageDo = Docx.DOCX_PAGE_3;
                    }
                    for (String s1 : pageDo) {
                        lineXml.append(StrUtil.replace(new StringBuilder(page), "<w:t>" + Docx.DOCX_PAGE_TEXT + "</w:t>", s1));
                    }
                }else if(s.length() > 4 && s.substring(0, 4).equals(Docx.DOCX_TABLE_IMG_KEY)){
                    int id = Integer.parseInt(s.substring(4));
                    Map<String, String> cssMap = new WordFileUtilImpl().imgCssDo(imgCss.get(id));
                    //经过清洗文档定位或者锚点定位一定是
                    if(Css.TRUE.equals(cssMap.get(Css.NEW_LINE))) {
                        //如果图片单独显示在一行，先把上一行的给解析完，留出地方储存这一行
                        if (!"".equals(lineXml.toString())) {
                            //行文本合成
                            xml.append(new WordFileUtilImpl().wordLineDo(lineXml, new StringBuilder()));
                            lineXml = new StringBuilder();
                        }
                    }
                    //文档定位：直接插到文档开头
                    if (Css.POSITION_FIXED.equals(cssMap.get(Css.POSITION))) {
                        xml.insert(0, new WordFileUtilImpl().wordImgDo(img.get(id), cssMap, paperWidth, paddingLeft,
                                paddingRight, paperHeight, paperHead, paperBottom, tempPath, es));
                    }
                    //锚点定位：直接插入到文档当前最后
                    if (Css.POSITION_ABSOLUTE.equals(cssMap.get(Css.POSITION))) {
                        xml.insert(0, new WordFileUtilImpl().wordImgDo(img.get(id), cssMap, paperWidth, paddingLeft,
                                paddingRight, paperHeight, paperHead, paperBottom, tempPath, es));
                        //锚点定位用户是否单独一行其实没有影响
                    }
                    //流定位：直接插入行最后
                    if(Css.POSITION_STATIC.equals(cssMap.get(Css.POSITION))){
                        String lineFlag = cssMap.get(Css.NEW_LINE);
                        lineXml.append(new WordFileUtilImpl().wordImgDo(img.get(id), cssMap, paperWidth, paddingLeft,
                                paddingRight, paperHeight, paperHead, paperBottom, tempPath, es));
                        //如果希望单独一行，则解析单独一行
                        if(Css.TRUE.equals(lineFlag)){
                            xml.append(new WordFileUtilImpl().wordLineDo(lineXml, new StringBuilder()));
                            lineXml = new StringBuilder();
                        }
                    }
                }else{
                    //就是普通的文本
                    lineXml.append(new WordFileUtilImpl().wordTestDo(s, css.toString(), es));
                }
            }else if(s.length() > 0){
                //行文本合成
                xml.append(new WordFileUtilImpl().wordLineDo(lineXml, contentCss));
                lineXml = new StringBuilder();
            }
        }
        //返回之前查看行缓存
        if(!"".equals(lineXml.toString())){
            //行文本合成
            xml.append(new WordFileUtilImpl().wordLineDo(lineXml, new StringBuilder()));
        }
        return xml.toString();
    }

    /**
     * 根据页眉页脚有关信息添加页眉页脚文件，并且返回需要添加在w:sectPr内的文本
     * @param paperHeadContext 页眉拼合的字符串
     * @param paperHeadContextCss 页眉整体样式
     * @param paperHeadImg 页眉图片列表
     * @param paperHeadImgCss 页眉图片的样式表
     * @param paperHeadText 页眉顶部的样式化文字
     * @param paperHeadCss 页眉顶部的样式化文字css
     * @param paperFootContext 页脚拼合字符串
     * @param paperFootContextCss 页脚整体样式
     * @param paperFootImg 页脚图片
     * @param paperFootImgCss 页脚的图片样式
     * @param paperFootText 页脚文字
     * @param paperFootCss 页脚css
     * @return 返回需要添加到wordXml中的xml信息
     */
    public StringBuilder wordPaperOutDo(String tempPath, StringBuilder paperHeadContext, StringBuilder paperHeadContextCss,
                                        List<String> paperHeadImg, List<String> paperHeadImgCss, List<String> paperHeadText,
                                        List<String> paperHeadCss, StringBuilder paperFootContext, StringBuilder paperFootContextCss,
                                        List<String> paperFootImg, List<String> paperFootImgCss, List<String> paperFootText,
                                        List<String> paperFootCss, int paperWidth, int paddingLeft, int paddingRight,
                                        int paperHeight, int paperHead, int paperBottom,  ExecutorService es){
        String uri = new IoUtil().getUri();
        String s = File.separator;
        IoUtil ioUtil = new IoUtil();
        //改content-type
        ioUtil.contentAddPaperOut(tempPath);
        //改setting
        ioUtil.copyPaperOutSetting(tempPath);
        //改_res/document
        int relNum = IoUtil.addPaperOutRelationship(tempPath);
        //添加style
        Map<String, String> headMap = null;
        if(paperHeadContextCss != null){
            headMap = StrUtil.cssToMap(paperHeadContextCss.toString());
        }
        Map<String, String> footMap = null;
        if(paperHeadContextCss != null){
            footMap = StrUtil.cssToMap(paperFootContextCss.toString());
        }
        StringBuilder styleXml = new StrUtil().paperOutGetStyle(headMap, footMap);
        IoUtil.paperOutAddStyle(styleXml, tempPath);
        //先把两个固定文件拷贝进去
        ioUtil.paperOutAddFile(tempPath, uri);
        //生成页眉拷贝进去
        StringBuilder header = new StringBuilder(textToXml(paperHeadContext, paperHeadContextCss, paperHeadImg, paperHeadImgCss, paperHeadText,
                paperHeadCss, tempPath, es, paperWidth, paddingLeft, paddingRight, paperHeight, paperHead, paperBottom));
        StrUtil.replace(header, "</w:pPr>", "##########");
        StrUtil.replace(header, "##########", "<w:pStyle w:val=\"a11\"/></w:pPr>");
        StringBuilder headerSb = new StringBuilder(ioUtil.readFileJar(s + "out" + s + "header.ppo").toString() + header + "</w:hdr>");

        IoUtil.writeTextToFile(tempPath + s + "word" + s + "header1.xml", headerSb);
        StringBuilder foot = new StringBuilder(textToXml(paperFootContext, paperFootContextCss, paperFootImg, paperFootImgCss, paperFootText,
                paperFootCss, tempPath, es, paperWidth, paddingLeft, paddingRight, paperHeight, paperHead, paperBottom));

        StrUtil.replace(foot, "</w:pPr>", "##########");
        StrUtil.replace(foot, "##########", "<w:pStyle w:val=\"a12\"/></w:pPr>");

        StringBuilder footSb = new StringBuilder(ioUtil.readFileJar(s + "out" + s + "footer.ppo").toString() + foot + "</w:ftr>");
        IoUtil.writeTextToFile(tempPath + s + "word" + s + "footer1.xml", footSb);
        //返回页面信息
        return new StringBuilder("<w:headerReference w:type=\"default\" r:id=\"rId" + relNum + "\"/><w:footerReference w:type=\"default\" r:id=\"rId" +
                (relNum + 1) + "\"/>");
    }

    /**
     * 生成纸张信息
     * @param paperWidth 纸张宽度
     * @param paperHeight 纸张高度
     * @param paddingTop 纸张上边距
     * @param paddingRight 纸张右边距
     * @param paddingBottom 纸张下边距
     * @param paddingLeft 纸张左边距
     * @param paperHeader 纸张顶部留空间
     * @param paperFooter 纸张底部留空间
     * @param cols 分栏
     * @param landscape 纸张是否需要横向
     * @return 返回xml
     */
    public StringBuilder wordPaperDo(int paperWidth, int paperHeight, int paddingTop, int paddingRight, int paddingBottom,
                                     int paddingLeft, int paperHeader, int paperFooter, StringBuilder cols, boolean landscape,
                                     String paperOut) {
        StringBuilder xml = new StringBuilder();
        xml.append("<w:sectPr w:rsidR=\"00AD336E\" w:rsidSect=\"00AD336E\">");
        xml.append(paperOut);
        xml.append("<w:pgSz w:w=\"");
        if(!landscape){
            xml.append(addWordPaperDo(paperWidth, paperHeight, paddingTop, paddingRight, paddingBottom, paddingLeft));
        }else{
            xml.append(addWordPaperDo(paperHeight, paperWidth, paddingRight, paddingBottom, paddingLeft, paddingTop));
        }
        xml.append("\" w:header=\"");
        xml.append(paperHeader);
        xml.append("\" w:footer=\"");
        xml.append(paperFooter);
        xml.append("\" w:gutter=\"0\"/>");
        xml.append(cols);
        xml.append("<w:docGrid w:type=\"lines\" w:linePitch=\"312\"/></w:sectPr>");
        return xml;
    }

    private static StringBuilder addWordPaperDo(int len1, int len2, int len3, int len4, int len5, int len6){
        StringBuilder xml = new StringBuilder();
        xml.append(len1);
        xml.append("\" w:h=\"");
        xml.append(len2);
        xml.append("\"/><w:pgMar w:top=\"");
        xml.append(len3);
        xml.append("\" w:right=\"");
        xml.append(len4);
        xml.append("\" w:bottom=\"");
        xml.append(len5);
        xml.append("\" w:left=\"");
        xml.append(len6);
        return xml;
    }

    /**
     * 生成文件
     * @param wordCore 文件信息
     * @param tempPath 临时储存目录
     * @param wordXml word生成的xml
     */
    public void wordCreate(WordCore wordCore, String tempPath, StringBuilder wordXml, String fileWay){
        String s = File.separator;
        //拼装document.xml
        wordXml.insert(0, Docx.DOC_HEAD);
        wordXml.append(Docx.DOC_FOOT);
        //拼装core
        StringBuilder core = StrUtil.coreDo(wordCore);
        //删除[Content_Types].xml注释
        IoUtil.contentCreate(tempPath);
        //删除word/res注释
        IoUtil.wordResCreate(tempPath);
        //删除word/style注释
        IoUtil.wordStyleCreate(tempPath);
        IoUtil.writeTextToFile(tempPath + s + "docProps" + s + "core.xml", core);
        IoUtil.writeTextToFile(tempPath + s + "word" + s + "document.xml", wordXml);
        IoUtil.zipMultiFile(tempPath,fileWay);
        IoUtil.deleteFile(new File(tempPath));
    }
}
