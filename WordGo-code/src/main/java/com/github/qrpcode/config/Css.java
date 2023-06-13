package com.github.qrpcode.config;

/**
 * 这里面就是储存一些常量的，一个字母为了方便调用
 * 不要删除或者修改，否则可能生成的word打不开
 * @author QiRuipeng
 */
public final class Css {
    public static final String PERCENT = "%";
    public static final String PX = "px";
    public static final String PT = "pt";
    public static final String AT = "@";

    //***********   以下是一些通用的属性   ***********

    public static final String BACKGROUND_COLOR = "background-color";
    public static final String COLOR = "color";
    public static final String FONT_FAMILY = "font-family";
    public static final String FONT_SIZE = "font-size";
    public static final String FONT_WEIGHT = "font-weight";
    public static final String FONT_STYLE = "font-style";
    /**
     * 着重号属性，当且仅当 font-emphasis:point 显示重点号
     */
    public static final String FONT_EMPHASIS = "font-emphasis";
    public static final String FONT_SCALE = "font-scale";
    public static final String TEXT_DECORATION = "text-decoration";
    /**
     * 下划线颜色属性，支持16位颜色，可以和文字颜色不一样
     * 只有设置了下划线样式时，这个属性才会有效果。
     */
    public static final String TEXT_DECORATION_COLOR = "text-decoration-color";
    public static final String TEXT_ALIGN = "text-align";
    public static final String LINE_HEIGHT = "line-height";
    public static final String TEXT_INDENT = "text-indent";
    public static final String NEW_LINE = "new-line";
    public static final String POSITION = "position";
    public static final String LEFT = "left";
    public static final String TOP = "top";
    public static final String WIDTH = "width";
    public static final String HEIGHT = "height";
    public static final String MARGIN = "margin";
    public static final String MARGIN_LEFT = "margin-left";
    public static final String MARGIN_RIGHT = "margin-right";
    public static final String MARGIN_TOP = "margin-top";
    public static final String MARGIN_BOTTOM = "margin-bottom";

    //***********   以下是一些通用属性的标记值   ***********

    public static final String A_BACKGROUND_COLOR_A = "@background-color@";
    public static final String A_COLOR_A = "@color@";
    public static final String A_FONT_FAMILY_A = "@font-family@";
    public static final String A_FONT_SIZE_A = "@font-size@";
    public static final String A_FONT_WEIGHT_A = "@font-weight@";
    public static final String A_FONT_STYLE_A = "@font-style@";
    public static final String A_FONT_EMPHASIS_A = "@font-emphasis@";
    public static final String A_FONT_SCALE_A = "@font-scale@";
    public static final String A_TEXT_DECORATION_A = "@text-decoration@";
    public static final String A_TEXT_ALIGN_A = "@text-align@";
    public static final String A_LINE_HEIGHT_A = "@line-height@";
    public static final String A_TEXT_INDENT_A = "@text-indent@";


    //***********   以下是图片的专属属性或固定值   ***********

    // box-shadow 相关（阴影）

    public static final String BOX_SHADOW = "box-shadow";
    public static final String A_BOX_SHADOW_A = "@box-shadow@";
    public static final String OUT_BOTTOM_RIGHT = "outbottomright";
    public static final String BOX_SHADOW_OUT_BOTTOM_RIGHT = "<a:outerShdw blurRad=\"50800\" dist=\"38100\" dir=\"2700000\" algn=\"tl\" rotWithShape=\"0\"><a:prstClr val=\"black\"><a:alpha val=\"40000\"/></a:prstClr></a:outerShdw>";
    public static final int BOX_SHADOW_OUT_BOTTOM_RIGHT_MARGIN_LEFT = 38100;
    public static final int BOX_SHADOW_OUT_BOTTOM_RIGHT_MARGIN_TOP = 38100;
    public static final int BOX_SHADOW_OUT_BOTTOM_RIGHT_MARGIN_RIGHT = 95250;
    public static final int BOX_SHADOW_OUT_BOTTOM_RIGHT_MARGIN_BOTTOM = 95250;
    public static final String OUT_BOTTOM = "outbottom";
    public static final String BOX_SHADOW_OUT_BOTTOM = "<a:outerShdw blurRad=\"50800\" dist=\"38100\" dir=\"5400000\" algn=\"t\" rotWithShape=\"0\"><a:prstClr val=\"black\"><a:alpha val=\"40000\"/></a:prstClr></a:outerShdw>";
    public static final int BOX_SHADOW_OUT_BOTTOM_MARGIN_LEFT = 57150;
    public static final int BOX_SHADOW_OUT_BOTTOM_MARGIN_TOP = 19050;
    public static final int BOX_SHADOW_OUT_BOTTOM_MARGIN_RIGHT = 57150;
    public static final int BOX_SHADOW_OUT_BOTTOM_MARGIN_BOTTOM = 95250;
    public static final String OUT_BOTTOM_LEFT = "outbottomleft";
    public static final String BOX_SHADOW_OUT_BOTTOM_LEFT = "<a:outerShdw blurRad=\"50800\" dist=\"38100\" dir=\"8100000\" algn=\"tr\" rotWithShape=\"0\"><a:prstClr val=\"black\"><a:alpha val=\"40000\"/></a:prstClr></a:outerShdw>";
    public static final int BOX_SHADOW_OUT_BOTTOM_LEFT_MARGIN_LEFT = 76200;
    public static final int BOX_SHADOW_OUT_BOTTOM_LEFT_MARGIN_TOP = 38100;
    public static final int BOX_SHADOW_OUT_BOTTOM_LEFT_MARGIN_RIGHT = 38100;
    public static final int BOX_SHADOW_OUT_BOTTOM_LEFT_MARGIN_BOTTOM = 95250;
    public static final String OUT_RIGHT = "outright";
    public static final String BOX_SHADOW_OUT_RIGHT = "<a:outerShdw blurRad=\"50800\" dist=\"38100\" algn=\"l\" rotWithShape=\"0\"><a:prstClr val=\"black\"><a:alpha val=\"40000\"/></a:prstClr></a:outerShdw>";
    public static final int BOX_SHADOW_OUT_RIGHT_MARGIN_LEFT = 19050;
    public static final int BOX_SHADOW_OUT_RIGHT_MARGIN_TOP = 57150;
    public static final int BOX_SHADOW_OUT_RIGHT_MARGIN_RIGHT = 95250;
    public static final int BOX_SHADOW_OUT_RIGHT_MARGIN_BOTTOM = 57150;
    public static final String OUT_CENTER = "outcenter";
    public static final String BOX_SHADOW_OUT_CENTER = "<a:outerShdw blurRad=\"63500\" sx=\"102000\" sy=\"102000\" algn=\"ctr\" rotWithShape=\"0\"><a:prstClr val=\"black\"><a:alpha val=\"40000\"/></a:prstClr></a:outerShdw>";
    public static final int BOX_SHADOW_OUT_CENTER_MARGIN_LEFT = 95250;
    public static final int BOX_SHADOW_OUT_CENTER_MARGIN_TOP = 76200;
    public static final int BOX_SHADOW_OUT_CENTER_MARGIN_RIGHT = 76200;
    public static final int BOX_SHADOW_OUT_CENTER_MARGIN_BOTTOM = 76200;
    public static final String OUT_LEFT = "outleft";
    public static final String BOX_SHADOW_OUT_LEFT = "<a:outerShdw blurRad=\"50800\" dist=\"38100\" dir=\"10800000\" algn=\"r\" rotWithShape=\"0\"><a:prstClr val=\"black\"><a:alpha val=\"40000\"/></a:prstClr></a:outerShdw>";
    public static final int BOX_SHADOW_OUT_LEFT_MARGIN_LEFT = 95250;
    public static final int BOX_SHADOW_OUT_LEFT_MARGIN_TOP = 57150;
    public static final int BOX_SHADOW_OUT_LEFT_MARGIN_RIGHT = 19050;
    public static final int BOX_SHADOW_OUT_LEFT_MARGIN_BOTTOM = 57150;
    public static final String OUT_TOP_RIGHT = "outtopright";
    public static final String BOX_SHADOW_OUT_TOP_RIGHT = "<a:outerShdw blurRad=\"50800\" dist=\"38100\" dir=\"18900000\" algn=\"bl\" rotWithShape=\"0\"><a:prstClr val=\"black\"><a:alpha val=\"40000\"/></a:prstClr></a:outerShdw>";
    public static final int BOX_SHADOW_OUT_TOP_RIGHT_MARGIN_LEFT = 38100;
    public static final int BOX_SHADOW_OUT_TOP_RIGHT_MARGIN_TOP = 95250;
    public static final int BOX_SHADOW_OUT_TOP_RIGHT_MARGIN_RIGHT = 95250;
    public static final int BOX_SHADOW_OUT_TOP_RIGHT_MARGIN_BOTTOM = 38100;
    public static final String OUT_TOP = "outtop";
    public static final String BOX_SHADOW_OUT_TOP = "<a:outerShdw blurRad=\"50800\" dist=\"38100\" dir=\"16200000\" rotWithShape=\"0\"><a:prstClr val=\"black\"><a:alpha val=\"40000\"/></a:prstClr></a:outerShdw>";
    public static final int BOX_SHADOW_OUT_TOP_MARGIN_LEFT = 57150;
    public static final int BOX_SHADOW_OUT_TOP_MARGIN_TOP = 95250;
    public static final int BOX_SHADOW_OUT_TOP_MARGIN_RIGHT = 57150;
    public static final int BOX_SHADOW_OUT_TOP_MARGIN_BOTTOM = 19050;
    public static final String OUT_TOP_LEFT = "outtopleft";
    public static final String BOX_SHADOW_OUT_TOP_LEFT = "<a:outerShdw blurRad=\"50800\" dist=\"38100\" dir=\"13500000\" algn=\"br\" rotWithShape=\"0\"><a:prstClr val=\"black\"><a:alpha val=\"40000\"/></a:prstClr></a:outerShdw>";
    public static final int BOX_SHADOW_OUT_TOP_LEFT_MARGIN_LEFT = 76200;
    public static final int BOX_SHADOW_OUT_TOP_LEFT_MARGIN_TOP = 95250;
    public static final int BOX_SHADOW_OUT_TOP_LEFT_MARGIN_RIGHT = 38100;
    public static final int BOX_SHADOW_OUT_TOP_LEFT_MARGIN_BOTTOM = 38100;
    public static final String PERSPECTIVE_TOP_LEFT = "perspectivetopleft";
    public static final String BOX_SHADOW_PERSPECTIVE_TOP_LEFT = "<a:outerShdw blurRad=\"76200\" dir=\"13500000\" sy=\"23000\" kx=\"1200000\" algn=\"br\" rotWithShape=\"0\"><a:prstClr val=\"black\"><a:alpha val=\"20000\"/></a:prstClr></a:outerShdw>";
    public static final int BOX_SHADOW_PERSPECTIVE_TOP_LEFT_MARGIN_LEFT = 495300;
    public static final int BOX_SHADOW_PERSPECTIVE_TOP_LEFT_MARGIN_TOP = 0;
    public static final int BOX_SHADOW_PERSPECTIVE_TOP_LEFT_MARGIN_RIGHT = 38100;
    public static final int BOX_SHADOW_PERSPECTIVE_TOP_LEFT_MARGIN_BOTTOM = 76200;
    public static final String PERSPECTIVE_TOP_RIGHT = "perspectivetopright";
    public static final String BOX_SHADOW_PERSPECTIVE_TOP_RIGHT = "<a:outerShdw blurRad=\"76200\" dir=\"18900000\" sy=\"23000\" kx=\"-1200000\" algn=\"bl\" rotWithShape=\"0\"><a:prstClr val=\"black\"><a:alpha val=\"20000\"/></a:prstClr></a:outerShdw>";
    public static final int BOX_SHADOW_PERSPECTIVE_TOP_RIGHT_MARGIN_LEFT = 38100;
    public static final int BOX_SHADOW_PERSPECTIVE_TOP_RIGHT_MARGIN_TOP = 0;
    public static final int BOX_SHADOW_PERSPECTIVE_TOP_RIGHT_MARGIN_RIGHT = 476250;
    public static final int BOX_SHADOW_PERSPECTIVE_TOP_RIGHT_MARGIN_BOTTOM = 76200;
    public static final String PERSPECTIVE_BOTTOM = "perspectivebottom";
    public static final String BOX_SHADOW_PERSPECTIVE_BOTTOM = "<a:outerShdw blurRad=\"152400\" dist=\"317500\" dir=\"5400000\" sx=\"90000\" sy=\"-19000\" rotWithShape=\"0\"><a:prstClr val=\"black\"><a:alpha val=\"15000\"/></a:prstClr></a:outerShdw>";
    public static final int BOX_SHADOW_PERSPECTIVE_BOTTOM_MARGIN_LEFT = 38100;
    public static final int BOX_SHADOW_PERSPECTIVE_BOTTOM_MARGIN_TOP = 0;
    public static final int BOX_SHADOW_PERSPECTIVE_BOTTOM_MARGIN_RIGHT = 38100;
    public static final int BOX_SHADOW_PERSPECTIVE_BOTTOM_MARGIN_BOTTOM = 685800;
    public static final String PERSPECTIVE_BOTTOM_LEFT = "perspectivebottomleft";
    public static final String BOX_SHADOW_PERSPECTIVE_BOTTOM_LEFT = "<a:outerShdw blurRad=\"76200\" dist=\"12700\" dir=\"8100000\" sy=\"-23000\" kx=\"800400\" algn=\"br\" rotWithShape=\"0\"><a:prstClr val=\"black\"><a:alpha val=\"20000\"/></a:prstClr></a:outerShdw>";
    public static final int BOX_SHADOW_PERSPECTIVE_BOTTOM_LEFT_MARGIN_LEFT = 361950;
    public static final int BOX_SHADOW_PERSPECTIVE_BOTTOM_LEFT_MARGIN_TOP = 0;
    public static final int BOX_SHADOW_PERSPECTIVE_BOTTOM_LEFT_MARGIN_RIGHT = 38100;
    public static final int BOX_SHADOW_PERSPECTIVE_BOTTOM_LEFT_MARGIN_BOTTOM = 381000;
    public static final String PERSPECTIVE_BOTTOM_RIGHT = "perspectivebottomright";
    public static final String BOX_SHADOW_PERSPECTIVE_BOTTOM_RIGHT = "<a:outerShdw blurRad=\"76200\" dist=\"12700\" dir=\"2700000\" sy=\"-23000\" kx=\"-800400\" algn=\"bl\" rotWithShape=\"0\"><a:prstClr val=\"black\"><a:alpha val=\"20000\"/></a:prstClr></a:outerShdw>";
    public static final int BOX_SHADOW_PERSPECTIVE_BOTTOM_RIGHT_MARGIN_LEFT = 38100;
    public static final int BOX_SHADOW_PERSPECTIVE_BOTTOM_RIGHT_MARGIN_TOP = 0;
    public static final int BOX_SHADOW_PERSPECTIVE_BOTTOM_RIGHT_MARGIN_RIGHT = 361950;
    public static final int BOX_SHADOW_PERSPECTIVE_BOTTOM_RIGHT_MARGIN_BOTTOM = 381000;
    public static final String IN_TOP_LEFT = "intopleft";
    public static final String BOX_SHADOW_IN_TOP_LEFT = "<a:innerShdw blurRad=\"63500\" dist=\"50800\" dir=\"13500000\"><a:prstClr val=\"black\"><a:alpha val=\"50000\"/></a:prstClr></a:innerShdw>";
    public static final String IN_TOP = "intop";
    public static final String BOX_SHADOW_IN_TOP = "<a:innerShdw blurRad=\"63500\" dist=\"50800\" dir=\"16200000\"><a:prstClr val=\"black\"><a:alpha val=\"50000\"/></a:prstClr></a:innerShdw>";
    public static final String IN_TOP_RIGHT = "intopright";
    public static final String BOX_SHADOW_IN_TOP_RIGHT = "<a:innerShdw blurRad=\"63500\" dist=\"50800\" dir=\"18900000\"><a:prstClr val=\"black\"><a:alpha val=\"50000\"/></a:prstClr></a:innerShdw>";
    public static final String IN_LEFT = "inleft";
    public static final String BOX_SHADOW_IN_LEFT = "<a:innerShdw blurRad=\"63500\" dist=\"50800\" dir=\"10800000\"><a:prstClr val=\"black\"><a:alpha val=\"50000\"/></a:prstClr></a:innerShdw>";
    public static final String IN_CENTER = "incenter";
    public static final String BOX_SHADOW_IN_CENTER = "<a:innerShdw blurRad=\"114300\"><a:prstClr val=\"black\"/></a:innerShdw>";
    public static final String IN_RIGHT = "inright";
    public static final String BOX_SHADOW_IN_RIGHT = "<a:innerShdw blurRad=\"63500\" dist=\"50800\"><a:prstClr val=\"black\"><a:alpha val=\"50000\"/></a:prstClr></a:innerShdw>";
    public static final String IN_BOTTOM_LEFT = "inbottomleft";
    public static final String BOX_SHADOW_IN_BOTTOM_LEFT = "<a:innerShdw blurRad=\"63500\" dist=\"50800\" dir=\"8100000\"><a:prstClr val=\"black\"><a:alpha val=\"50000\"/></a:prstClr></a:innerShdw>";
    public static final String IN_BOTTOM = "inbottom";
    public static final String BOX_SHADOW_IN_BOTTOM = "<a:innerShdw blurRad=\"63500\" dist=\"50800\" dir=\"5400000\"><a:prstClr val=\"black\"><a:alpha val=\"50000\"/></a:prstClr></a:innerShdw>";
    public static final String IN_BOTTOM_RIGHT = "inbottomright";
    public static final String BOX_SHADOW_IN_BOTTOM_RIGHT = "<a:innerShdw blurRad=\"63500\" dist=\"50800\" dir=\"2700000\"><a:prstClr val=\"black\"><a:alpha val=\"50000\"/></a:prstClr></a:innerShdw>";

    // reflection 属性（映像）

    public static final String REFLECTION = "reflection";
    public static final String A_REFLECTION_A = "@reflection@";
    public static final String NEAR_MIN = "nearmin";
    public static final String REFLECTION_NEAR_MIN = "<a:reflection blurRad=\"6350\" stA=\"52000\" endA=\"300\" endPos=\"35000\" dir=\"5400000\" sy=\"-100000\" algn=\"bl\" rotWithShape=\"0\"/>";
    public static final int REFLECTION_NEAR_MIN_LEFT = 19050;
    public static final int REFLECTION_NEAR_MIN_TOP = 0;
    public static final int REFLECTION_NEAR_MIN_RIGHT = 19050;
    public static final int REFLECTION_NEAR_MIN_BOTTOM = 457200;
    public static final String NEAR = "near";
    public static final String REFLECTION_NEAR = "<a:reflection blurRad=\"6350\" stA=\"50000\" endA=\"300\" endPos=\"55000\" dir=\"5400000\" sy=\"-100000\" algn=\"bl\" rotWithShape=\"0\"/>";
    public static final int REFLECTION_NEAR_LEFT = 19050;
    public static final int REFLECTION_NEAR_TOP = 0;
    public static final int REFLECTION_NEAR_RIGHT = 19050;
    public static final int REFLECTION_NEAR_BOTTOM = 704850;
    public static final String NEAR_MAX = "nearmax";
    public static final String REFLECTION_NEAR_MAX = "<a:reflection blurRad=\"6350\" stA=\"50000\" endA=\"300\" endPos=\"90000\" dir=\"5400000\" sy=\"-100000\" algn=\"bl\" rotWithShape=\"0\"/>";
    public static final int REFLECTION_NEAR_MAX_LEFT = 19050;
    public static final int REFLECTION_NEAR_MAX_TOP = 0;
    public static final int REFLECTION_NEAR_MAX_RIGHT = 19050;
    public static final int REFLECTION_NEAR_MAX_BOTTOM = 1143000;
    public static final String MEDIUM_MIN = "mediummin";
    public static final String REFLECTION_MEDIUM_MIN = "<a:reflection blurRad=\"6350\" stA=\"50000\" endA=\"300\" endPos=\"38500\" dist=\"50800\" dir=\"5400000\" sy=\"-100000\" algn=\"bl\" rotWithShape=\"0\"/>";
    public static final int REFLECTION_MEDIUM_MIN_LEFT = 19050;
    public static final int REFLECTION_MEDIUM_MIN_TOP = 0;
    public static final int REFLECTION_MEDIUM_MIN_RIGHT = 19050;
    public static final int REFLECTION_MEDIUM_MIN_BOTTOM = 1143000;
    public static final String MEDIUM = "medium";
    public static final String REFLECTION_MEDIUM = "<a:reflection blurRad=\"6350\" stA=\"50000\" endA=\"300\" endPos=\"55500\" dist=\"50800\" dir=\"5400000\" sy=\"-100000\" algn=\"bl\" rotWithShape=\"0\"/>";
    public static final int REFLECTION_MEDIUM_LEFT = 19050;
    public static final int REFLECTION_MEDIUM_TOP = 0;
    public static final int REFLECTION_MEDIUM_RIGHT = 19050;
    public static final int REFLECTION_MEDIUM_BOTTOM = 762000;
    public static final String MEDIUM_MAX = "mediummax";
    public static final String REFLECTION_MEDIUM_MAX = "<a:reflection blurRad=\"6350\" stA=\"50000\" endA=\"300\" endPos=\"90000\" dist=\"50800\" dir=\"5400000\" sy=\"-100000\" algn=\"bl\" rotWithShape=\"0\"/>";
    public static final int REFLECTION_MEDIUM_MAX_LEFT = 19050;
    public static final int REFLECTION_MEDIUM_MAX_TOP = 0;
    public static final int REFLECTION_MEDIUM_MAX_RIGHT = 19050;
    public static final int REFLECTION_MEDIUM_MAX_BOTTOM = 1143000;
    public static final String FAR_MIN = "farmin";
    public static final String REFLECTION_FAR_MIN = "<a:reflection blurRad=\"6350\" stA=\"50000\" endA=\"275\" endPos=\"40000\" dist=\"101600\" dir=\"5400000\" sy=\"-100000\" algn=\"bl\" rotWithShape=\"0\"/>";
    public static final int REFLECTION_FAR_MIN_LEFT = 19050;
    public static final int REFLECTION_FAR_MIN_TOP = 0;
    public static final int REFLECTION_FAR_MIN_RIGHT = 19050;
    public static final int REFLECTION_FAR_MIN_BOTTOM = 590550;
    public static final String FAR = "far";
    public static final String REFLECTION_FAR = "<a:reflection blurRad=\"6350\" stA=\"50000\" endA=\"300\" endPos=\"55500\" dist=\"101600\" dir=\"5400000\" sy=\"-100000\" algn=\"bl\" rotWithShape=\"0\"/>";
    public static final int REFLECTION_FAR_LEFT = 19050;
    public static final int REFLECTION_FAR_TOP = 0;
    public static final int REFLECTION_FAR_RIGHT = 19050;
    public static final int REFLECTION_FAR_BOTTOM = 800100;
    public static final String FAR_MAX = "farmin";
    public static final String REFLECTION_FAR_MAX = "<a:reflection blurRad=\"6350\" stA=\"50000\" endA=\"295\" endPos=\"92000\" dist=\"101600\" dir=\"5400000\" sy=\"-100000\" algn=\"bl\" rotWithShape=\"0\"/>";
    public static final int REFLECTION_FAR_MAX_LEFT = 19050;
    public static final int REFLECTION_FAR_MAX_TOP = 0;
    public static final int REFLECTION_FAR_MAX_RIGHT = 19050;
    public static final int REFLECTION_FAR_MAX_BOTTOM = 1257300;

    //柔化边缘有关

    public static final String SOFT_EDGE = "soft-edge";
    public static final String A_SOFT_EDGE_A = "@soft-edge@";
    public static final int SOFT_EDGE_TIMES = 12700;

    //边框有关

    public static final String BORDER = "border";
    public static final String A_BORDER_A = "@border@";
    public static final String BORDER_A3 = "a3";
    public static final int BORDER_TIMES = 12700;
    public static final String[] BORDER_TYPE = {"solid", "dotted", "sysDash", "dashed", "dashDot", "lgDash", "lgDashDot", "lgDashDotDot"};
    public static final String[] BORDER_TYPE_XML = {"", "<a:prstDash val=\"sysDot\"/>", "<a:prstDash val=\"sysDash\"/>",
            "<a:prstDash val=\"dash\"/>", "<a:prstDash val=\"dashDot\"/>", "<a:prstDash val=\"lgDash\"/>", "<a:prstDash val=\"lgDashDot\"/>",
            "<a:prstDash val=\"lgDashDotDot\"/>"};
    public static final String BORDER_WIDTH = "border-width";
    public static final String BORDER_STYLE = "border-style";
    public static final String BORDER_COLOR = "border-color";

    //***********   以下是一些表格中的参数   ***********

    public static final String BORDER_TOP = "border-top";
    public static final String BORDER_BOTTOM = "border-bottom";
    public static final String BORDER_LEFT = "border-left";
    public static final String BORDER_RIGHT = "border-right";
    public static final String BORDER_INSIDE = "border-inside";
    public static final String BORDER_INSIDE_H = "border-insideH";
    public static final String BORDER_INSIDE_V = "border-insideV";
    public static final String[] BORDER_TABLE = {"nil", "single", "dotted", "dashSmallGap", "dashed", "dotDash",
            "dotDotDash", "double", "triple", "thinThickSmallGap", "thickThinSmallGap", "thinThickThinSmallGap",
            "thinThickLargeGap", "thickThinLargeGap", "thinThickThinLargeGap", "wave", "doubleWave", "dashDotStroked",
            "threeDEmboss", "threeDEngrave"};
    public static final String[] BORDER_TABLE_LOW = {"nil", "single", "dotted", "dashsmallgap", "dashed", "dotdash",
            "dotdotdash", "double", "triple", "thinthicksmallgap", "thickthinsmallgap", "thinthickthinsmallgap",
            "thinthicklargegap", "thickthinlargegap", "thinthickthinlargegap", "wave", "doublewave", "dashdotstroked",
            "threedemboss", "threedengrave"};

    public static final String ROW_HEIGHT = "row-height";
    public static final String COLUMN_WIDTH = "column-width";
    public static final String TEMPLATE = "template";
    public static final String NORMAL = "normal";
    public static final String GIRD = "gird";
    public static final String LIST = "list";

    //***********   以下是一些通用属性的可赋予值   ***********

    public static final String TEXT_ALIGN_LEFT = "left";
    public static final String TEXT_ALIGN_RIGHT = "right";
    public static final String TEXT_ALIGN_CENTER = "center";

    public static final String POSITION_STATIC = "static";
    public static final String POSITION_FIXED = "fixed";
    public static final String POSITION_ABSOLUTE = "absolute";

    public static final String TRUE = "true";
    public static final String FALSE = "false";
}
