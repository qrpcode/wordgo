package com.github.qrpcode.util;

import com.github.qrpcode.domain.WordCore;
import com.github.qrpcode.domain.WordTable;

import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutorService;

/**
 * 初始生成一个word文件
 * @author QiRuipeng
 */
public interface WordFileUtil {
    /**
     * 初始化一个项目，其实就是把各种奇奇怪怪还没变化的固定文件拷贝到一个临时路径之后把临时路径返回回来
     * @return 返回临时储存工作路径
     */
    String startWordGo();

    /**
     * 对一串文本进行解析加工，主要就是解析css，推荐使用多线程哟
     * @param text 文本内容
     * @param css 文本对应的样式表文本，纯文本就行，解析成map是方法需要做的
     * @param es 线程池
     * @return 返回解析成docx文件存储形式的XML信息
     */
    StringBuilder wordTestDo(String text, String css, ExecutorService es);

    /**
     * 在word中切换到新的一页，不管现在这个页面写了多少内容
     * @return 返回解析成docx文件存储形式的XML信息
     */
    StringBuilder wordNewPageDo();

    /**
     * 这个是一个图片的预处理函数，解析加工图片之前必须先执行这个方法。
     * 为什么呢，因为他很需要知道这张图片的定位是什么样的，根据定位确定将这张图片的XML插入到什么位置，所以我们需要把结果返回回去
     * 等会解析加工的时候直接就传这个map集合了
     * 有必要提醒，这个方法应该具有自动补全是否需要独立成行和定位属性，即使用户没有定义
     *
     * @param css 图片对应的样式表文本
     * @return 返回解析后的map集合
     */
    Map<String, String> imgCssDo(String css);

    /**
     * 对一张图片进行解析加工
     * @param imgUri 图片的实际路径，只能是本地的图片，网络图片写个链接解析不出来的
     * @param cssMap css的Map集合
     * @param paperWidth 纸张的宽度
     * @param paddingLeft 纸张的左边距
     * @param paddingRight 纸张的右边距
     * @param paperHeight 纸张的高度
     * @param paperHead 纸张的顶边距
     * @param paperBottom 纸张的底边距
     * @param tempPath 文件临时储存目录
     * @param es 线程池
     * @return 返回解析成docx文件存储形式的XML信息
     */
    StringBuilder wordImgDo(String imgUri, Map<String, String> cssMap, int paperWidth, int paddingLeft, int paddingRight,
                            int paperHeight, int paperHead, int paperBottom, String tempPath, ExecutorService es);

    /**
     * 对一行文本进行解析加工，这一行应该是里面文本图片啥的都已经解析过了，现在只需要套在一个行信息里面，而不是从0开始制作整个一行信息
     * @param lineXml 这行中文本的XML信息
     * @param css 整个一行对应的样式表文本
     * @return 返回整个一行docx文件存储形式的XML信息
     */
    StringBuilder wordLineDo(StringBuilder lineXml, StringBuilder css);

    /**
     * 解析一个表格
     * @param wordTable 表格类
     * @param tempPath 文件临时储存目录
     * @param es 线程池
     * @param paperHeight 纸张高度
     * @param paddingTop 纸张顶边距
     * @param paddingBottom 纸张下边距
     * @param paperHead 页眉留白高度
     * @param paperFoot 页脚留白高度
     * @param paperWidth 纸张宽度
     * @param paddingLeft 页面左边距
     * @param paddingRight 页面右边距
     * @return 返回解析成docx文件存储形式的XML信息
     */
    StringBuilder wordTableDo(WordTable wordTable, String tempPath, ExecutorService es,
                              int paperHeight, int paddingTop, int paddingBottom, int paperHead, int paperFoot,
                              int paperWidth, int paddingLeft, int paddingRight);

    /**
     * 对页眉页脚进行处理，在返回信息的同时还需要准备好页眉页脚信息文件写入到临时目录中
     * @param tempPath 文件临时储存目录
     * @param paperHeadContext 页眉信息文本
     * @param paperHeadContextCss 页眉样式表文本
     * @param paperHeadImg 页眉图片列表
     * @param paperHeadImgCss 页眉图片css列表
     * @param paperHeadText 页眉具有独特样式的文本列表
     * @param paperHeadCss 页眉具有独特样式的文本对应css列表
     * @param paperFootContext 页脚信息文本
     * @param paperFootContextCss 页脚样式表文本
     * @param paperFootImg 页脚图片列表
     * @param paperFootImgCss 页脚图片css列表
     * @param paperFootText 页脚具有独特样式的文本列表
     * @param paperFootCss 页脚具有独特样式的文本对应css列表
     * @param paperWidth 纸张宽度
     * @param paddingLeft 纸张左边距
     * @param paddingRight 纸张右边距
     * @param paperHeight 纸张上边距
     * @param paperHead 页眉留白高度
     * @param paperBottom 页脚留白高度
     * @param es 线程池
     * @return 返回解析成docx文件存储形式的XML信息
     */
    StringBuilder wordPaperOutDo(String tempPath, StringBuilder paperHeadContext, StringBuilder paperHeadContextCss, List<String> paperHeadImg,
                          List<String> paperHeadImgCss, List<String> paperHeadText, List<String> paperHeadCss,
                          StringBuilder paperFootContext, StringBuilder paperFootContextCss, List<String> paperFootImg,
                          List<String> paperFootImgCss, List<String> paperFootText, List<String> paperFootCss
                          , int paperWidth, int paddingLeft, int paddingRight, int paperHeight,
                          int paperHead, int paperBottom, ExecutorService es);

    /**
     * 解析页面纸张信息
     * @param paperWidth 纸张宽度
     * @param paperHeight 纸张高度
     * @param paddingTop 纸张上边距
     * @param paddingRight 纸张右边距
     * @param paddingBottom 纸张下边距
     * @param paddingLeft 纸张左边距
     * @param paperHeader 纸张上边距
     * @param paperFooter 纸张下边距
     * @param cols 分栏信息（分栏信息是已经解析过的xml直接拼装即可）
     * @param landscape 是否需要横向展示（就是word里面的纸张方向），一般来说都是false正常展示
     * @param paperOut 纸张页眉页脚信息（纸张页眉页脚信息是已经解析过的xml直接拼装即可）
     * @return 返回解析成docx文件存储形式的XML信息
     */
    StringBuilder wordPaperDo(int paperWidth, int paperHeight, int paddingTop, int paddingRight, int paddingBottom,
                       int paddingLeft, int paperHeader, int paperFooter, StringBuilder cols, boolean landscape,
                       String paperOut);

    /**
     * 最终生成一个
     * @param tempPath 文件临时储存目录
     * @param wordXml word的XML信息
     * @param fileWay 生成路径，需要包含文件名，文件名必须以.docx结尾才能打开
     * @param wordCore 文件用户信息
     */
    void wordCreate(WordCore wordCore, String tempPath, StringBuilder wordXml, String fileWay);
}
