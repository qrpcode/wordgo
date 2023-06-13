package com.github.qrpcode.config;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

/**
 * 储存一些常量
 * @author Qiruipeng
 */
public final class Docx {
    public static final int DOCX_EUM = 9525;
    public static final int DOCX_PROPORTION = 636;
    public static final double DOCX_PAPER = 56.69428;

    public static final String DOCX_PAPER_A3 = "a3";
    public static final String DOCX_PAPER_A4 = "a4";
    public static final String DOCX_PAPER_A5 = "a5";
    public static final String DOCX_PAPER_B4_JIS = "b4jis";
    public static final String DOCX_PAPER_B5_JIS = "b5jis";
    public static final String DOCX_PAPER_TABLOID = "tabloid";
    public static final String DOCX_PAPER_STATEMENT = "statement";
    public static final String DOCX_PAPER_EXECUTIVE = "executive";
    public static final String DOCX_PAPER_LETTER = "letter";
    public static final String DOCX_PAPER_LAW = "legal";
    public static final String DOCX_PAPER_16CUT = "16cut";
    public static final String DOCX_PAPER_32CUT = "32cut";
    public static final String DOCX_PAPER_BIG32CUT = "big32cut";

    public static final String DOCX_AMP = "&";
    public static final String DOCX_AMP_CHOOSE = "&amp;";
    public static final String DOCX_LT = "<";
    public static final String DOCX_LT_CHOOSE = "&lt;";
    public static final String DOCX_GT = ">";
    public static final String DOCX_GT_CHOOSE = "&gt;";

    public static final String DOCX_FORMAT_JPG = "jpg";
    public static final String DOCX_FORMAT_PNG = "png";
    public static final String DOCX_FORMAT_GIF = "gif";
    public static final String DOCX_FORMAT_JPEG = "jpeg";

    public static final String DOCX_TABLE_KEY = "@GO#";
    public static final String DOCX_TABLE_CUT = "@@";
    public static final String DOCX_TABLE_HASH = "#";
    public static final String DOCX_TABLE_IMG_KEY = "@IMG";
    public static final String DOCX_TABLE_CSS_KEY = "@CSS";
    public static final String DOCX_TABLE_HEAD = "<w:tbl><w:tblPr><w:tblStyle w:val=\"@template@\"/><w:tblW w:w=\"0\" w:type=\"auto\"/>@text-align@@tblBorders@<w:tblLook w:val=\"04A0\" w:firstRow=\"1\" w:lastRow=\"0\" w:firstColumn=\"1\" w:lastColumn=\"0\" w:noHBand=\"0\" w:noVBand=\"1\"/></w:tblPr><w:tblGrid>@column-width@</w:tblGrid>";
    public static final String DOCX_TABLE_FOOT = "</w:tbl>";
    public static final String DOCX_TABLE_ROW_1 = "<w:tr w:rsidR=\"00AD336E\" w:rsidTr=\"00AD336E\">";
    public static final String DOCX_TABLE_ROW_2 = "</w:tr>";
    public static final String DOCX_TABLE_TRPR_1 = "<w:trPr>";
    public static final String DOCX_TABLE_ROW_WIDTH_1 = "<w:trHeight w:val=\"";
    public static final String DOCX_TABLE_ROW_WIDTH_2 = "\"/>";
    public static final String DOCX_TABLE_TRPR_2 = "</w:trPr>";
    public static final String DOCX_TABLE_COLUMN_WIDTH_1 = "<w:tc><w:tcPr><w:tcW w:w=\"";
    public static final String DOCX_TABLE_COLUMN_WIDTH_2 = "\" w:type=\"dxa\"/>";
    public static final String DOCX_TABLE_TR_2 = "</w:tc>";

    public static final String DOCX_CONTENT = "[Content_Types].xml";
    public static final String DOCX_CONTENT_TYPE_LEFT = "Extension=\"";
    public static final String DOCX_CONTENT_TYPE_RIGHT = "\"";
    public static final String DOCX_CONTENT_PIC_GET = "<!--picGet-->";
    public static final String DOCX_CONTENT_TYPE_ADD_HEAD = "<Default Extension=\"";
    public static final String DOCX_CONTENT_TYPE_ADD_BODY = "\" ContentType=\"image/";
    public static final String DOCX_CONTENT_TYPE_ADD_FOOT = "\"/>";

    public static final String DOCX_RELATIONSHIP = File.separator + "word" + File.separator + "_rels" + File.separator + "document.xml.rels";
    public static final String DOCX_RELATIONSHIP_GET = "<!--res-->";
    public static final String DOCX_RELATIONSHIP_LABEL = "<Relationship ";
    public static final String DOCX_RELATIONSHIP_TEXT_1 = "<Relationship Id=\"rId";
    public static final String DOCX_RELATIONSHIP_TEXT_2 = "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"media/image";
    public static final String DOCX_RELATIONSHIP_TEXT_3 = ".";
    public static final String DOCX_RELATIONSHIP_TEXT_4 = "\"/>";

    public static final String DOCX_NEW_LINE = "<w:p w:rsidR=\"00AD336E\" w:rsidRDefault=\"00AD336E\"><w:r><w:br w:type=\"page\"/></w:r></w:p><w:p w:rsidR=\"00AD336E\" w:rsidRDefault=\"00AD336E\"><w:pPr><w:spacing w:line=\"20\" w:lineRule=\"exact\"/></w:pPr></w:p>";
    public static final String DOCX_PAGE = "@PGE";
    public static final String DOCX_PAGE_TEXT = "666888";
    public static final String[] DOCX_PAGE_1 = {"<w:fldChar w:fldCharType=\"begin\"/>", "<w:instrText>PAGE   \\* MERGEFORMAT</w:instrText>", "<w:fldChar w:fldCharType=\"separate\"/>", "<w:t>1</w:t>", "<w:fldChar w:fldCharType=\"end\"/>"};
    public static final String[] DOCX_PAGE_2 = {"<w:t xml:space=\"preserve\"></w:t>", "<w:fldChar w:fldCharType=\"begin\"/>", "<w:instrText>PAGE  \\* Arabic  \\* MERGEFORMAT</w:instrText>", "<w:fldChar w:fldCharType=\"separate\"/>", "<w:t>1</w:t>", "<w:fldChar w:fldCharType=\"end\"/>", "<w:t xml:space=\"preserve\"> / </w:t>", "<w:fldChar w:fldCharType=\"begin\"/>", "<w:instrText>NUMPAGES  \\* Arabic  \\* MERGEFORMAT</w:instrText>", "<w:fldChar w:fldCharType=\"separate\"/>", "<w:t>1</w:t>", "<w:fldChar w:fldCharType=\"end\"/>"};
    public static final String[] DOCX_PAGE_3 = {"<w:t xml:space=\"preserve\">~ </w:t>", "<w:fldChar w:fldCharType=\"begin\"/>", "<w:instrText>PAGE \\* MERGEFORMAT</w:instrText>", "<w:fldChar w:fldCharType=\"separate\"/>", "<w:t>1</w:t>", "<w:fldChar w:fldCharType=\"end\"/>", "<w:t xml:space=\"preserve\"> ~</w:t>"};

    public static final String DOCX_STYLE = File.separator + "word" + File.separator + "styles.xml";
    public static final String DOCX_STYLE_HEAD = "<w:style w:type=\"table\" w:styleId=\"";
    public static final String DOCX_STYLE_FOOT = "\">";
    public static final String DOCX_STYLE_ADD = "<!--add-->";
    public static final String DOCX_OUT_HEAD_URI = File.separator + "out" + File.separator + "head.ppo";
    public static final String DOCX_OUT_HEAD_DIY_URI = File.separator + "out" + File.separator + "headdiy.ppo";
    public static final String DOCX_OUT_FOOT_URI = File.separator + "out" + File.separator + "foot.ppo";
    public static final String DOCX_OUT_FOOT_DIY_URI = File.separator + "out" + File.separator + "footdiy.ppo";

    public static final String DOC_HEAD = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<w:document xmlns:wpc=\"http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas\" xmlns:cx=\"http://schemas.microsoft.com/office/drawing/2014/chartex\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp14=\"http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\" xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" xmlns:w16se=\"http://schemas.microsoft.com/office/word/2015/wordml/symex\" xmlns:wpg=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\" xmlns:wpi=\"http://schemas.microsoft.com/office/word/2010/wordprocessingInk\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\" xmlns:wps=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\" mc:Ignorable=\"w14 w15 w16se wp14\"><w:body>";
    public static final String DOC_FOOT = "</w:body></w:document>";
    public static final String DOC_TEXT = "<w:r><w:rPr>@font-family@@font-scale@@font-size@@font-weight@@font-style@@background-color@@color@@text-decoration-color@@font-emphasis@</w:rPr><w:t>#text#</w:t></w:r>";
    public static final String DOC_LINE = "<w:p w:rsidR=\"000D4FB6\" w:rsidRDefault=\"00AD336E\"><w:pPr>@text-indent@@text-align@@line-height@</w:pPr>#lineXml#</w:p>";
    public static final String DOC_IMG_STATIC = "<w:r><w:rPr><w:rFonts w:hint=\"eastAsia\"/><w:noProof/></w:rPr><w:drawing><wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\"><wp:extent cx=\"@width@\" cy=\"@height@\"/><wp:effectExtent l=\"@margin-left@\" t=\"@margin-top@\" r=\"@margin-right@\" b=\"@margin-bottom@\"/><wp:docPr id=\"#imgId#\" name=\"@title@\"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/></wp:cNvGraphicFramePr><a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:nvPicPr><pic:cNvPr id=\"#imgId#\" name=\"@title@\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"rId#rid#\"><a:extLst><a:ext uri=\"{28A0092B-C50C-407E-A947-70E740481C1C}\"><a14:useLocalDpi xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" val=\"0\"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"@width@\" cy=\"@height@\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>@border@<a:effectLst>@box-shadow@@reflection@@glow@@rotate@@soft-edge@</a:effectLst></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";
    public static final String DOC_IMG_FIXED = "<w:p w:rsidR=\"00AD336E\" w:rsidRDefault=\"00AD336E\" w:rsidP=\"00AD336E\"><w:pPr><w:spacing w:line=\"20\" w:lineRule=\"exact\"/></w:pPr><w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:anchor distT=\"0\" distB=\"0\" distL=\"114300\" distR=\"114300\" simplePos=\"0\" relativeHeight=\"251660288\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\" wp14:anchorId=\"00F94124\" wp14:editId=\"00F94124\"><wp:simplePos x=\"0\" y=\"0\"/><wp:positionH relativeFrom=\"column\"><wp:posOffset>@left@</wp:posOffset></wp:positionH><wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>@top@</wp:posOffset></wp:positionV><wp:extent cx=\"@width@\" cy=\"@height@\"/><wp:effectExtent l=\"@margin-left@\" t=\"@margin-top@\" r=\"@margin-right@\" b=\"@margin-bottom@\"/><wp:wrapNone/><wp:docPr id=\"#imgId#\" name=\"@title@\"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/></wp:cNvGraphicFramePr><a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:nvPicPr><pic:cNvPr id=\"#imgId#\" name=\"@title@\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"rId#rid#\"><a:extLst><a:ext uri=\"{28A0092B-C50C-407E-A947-70E740481C1C}\"><a14:useLocalDpi xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" val=\"0\"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"@width@\" cy=\"@height@\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>@border@<a:effectLst>@box-shadow@@reflection@@glow@@rotate@@soft-edge@</a:effectLst></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>";
    public static final String DOC_IMG_ABSOLUTE = "<w:r><w:rPr><w:noProof/></w:rPr><w:drawing><wp:anchor distT=\"0\" distB=\"0\" distL=\"114300\" distR=\"114300\" simplePos=\"0\" relativeHeight=\"251660288\" behindDoc=\"0\" locked=\"0\" layoutInCell=\"1\" allowOverlap=\"1\" wp14:anchorId=\"00F94124\" wp14:editId=\"00F94124\"><wp:simplePos x=\"0\" y=\"0\"/><wp:positionH relativeFrom=\"column\"><wp:posOffset>@left@</wp:posOffset></wp:positionH><wp:positionV relativeFrom=\"paragraph\"><wp:posOffset>@top@</wp:posOffset></wp:positionV><wp:extent cx=\"@width@\" cy=\"@height@\"/><wp:effectExtent l=\"@margin-left@\" t=\"@margin-top@\" r=\"@margin-right@\" b=\"@margin-bottom@\"/><wp:wrapNone/><wp:docPr id=\"#imgId#\" name=\"@title@\"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/></wp:cNvGraphicFramePr><a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"><pic:nvPicPr><pic:cNvPr id=\"#imgId#\" name=\"@title@\"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed=\"rId#rid#\"><a:extLst><a:ext uri=\"{28A0092B-C50C-407E-A947-70E740481C1C}\"><a14:useLocalDpi xmlns:a14=\"http://schemas.microsoft.com/office/drawing/2010/main\" val=\"0\"/></a:ext></a:extLst></a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"@width@\" cy=\"@height@\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>@border@<a:effectLst>@box-shadow@@reflection@@glow@@rotate@@soft-edge@</a:effectLst></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>";
}
