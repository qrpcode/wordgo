package com.github.qrpcode.util;

import com.github.qrpcode.config.Docx;

import java.awt.image.BufferedImage;
import java.io.*;
import java.net.URL;
import java.util.Random;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

/**
 * 输出IO流
 * @author QiRuipeng
 */
public class IoUtil {

    private static final int  BUFFER_SIZE = 2 * 1024;

    String getUri(){
        //return this.getClass().getClassLoader().getResource("").getFile().substring(1);
        return this.getClass().getProtectionDomain().getCodeSource().getLocation().getFile().substring(1);
    }

    String buildUtil(String uri){
        String s = File.separator;
        String fileTemp = System.getProperty("user.dir") + s + "temp" + s + getRandomString(32);
        new File(fileTemp + s + "_rels").mkdirs();
        new File(fileTemp + s + "docProps").mkdirs();
        new File(fileTemp + s + "word" + s + "_rels").mkdirs();
        new File(fileTemp + s + "word" + s + "theme").mkdirs();
        //开始复制文件
        try {
            copyFileJar(s +"Content_Types.xml2", fileTemp + s + "[Content_Types].xml");
            copyFileJar(s +"rels" + s + "0.rels", fileTemp + s + "_rels" + s + ".rels");
            copyFileJar(s +"docProps" + s + "app.xml2", fileTemp + s + "docProps" + s + "app.xml");
            copyFileJar(s +"docProps" + s + "core.xml2", fileTemp + s + "docProps" + s + "core.xml");
            copyFileJar(s +"word" + s + "fontTable.xml2", fileTemp + s + "word" + s + "fontTable.xml");
            copyFileJar(s +"word" + s + "settings.xml2", fileTemp + s + "word" + s + "settings.xml");
            copyFileJar(s +"word" + s + "styles.xml2", fileTemp + s + "word" + s + "styles.xml");
            copyFileJar(s +"word" + s + "websettings.xml2", fileTemp + s + "word" + s + "webSettings.xml");
            copyFileJar(s +"word" + s + "_rels" + s + "document.xml.rels", fileTemp + s + "word" + s + "_rels" + s + "document.xml.rels");
            copyFileJar(s +"word" + s + "theme" + s + "theme1.xml2", fileTemp + s + "word" + s + "theme" + s + "theme1.xml");
            return fileTemp;
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("创建项目时读写临时文件失败！");
        }
    }

    private static String getCurrentDirPath() {
        URL url = IoUtil.class.getProtectionDomain().getCodeSource().getLocation();
        String path = url.getPath();
        if(path.startsWith("file:")) {
            path = path.replace("file:", "");
        }
        if(path.contains(".jar!/")) {
            path = path.substring(0, path.indexOf(".jar!/")+4);
        }

        File file = new File(path);
        path = file.getParentFile().getAbsolutePath();
        return path;
    }

    private static String getRandomString(int length){
        String str="abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
        Random random=new Random();
        StringBuffer sb=new StringBuffer();
        for(int i=0;i<length;i++){
            int number=random.nextInt(62);
            sb.append(str.charAt(number));
        }
        return sb.toString();
    }

    private static void copyFile(String copyFile, String newFile) throws Exception {
        FileInputStream fis = null;
        FileOutputStream fos = null;
        try{
            fis = new FileInputStream(copyFile);
            fos = new FileOutputStream(newFile, true);
            byte[] bytes = new byte[1024];
            int len = 0;
            while((len = fis.read(bytes)) != -1){
                fos.write(bytes, 0, len);
            }
        }catch (Exception e){
            e.printStackTrace();
        }finally {
            if(fos != null){
                fos.close();
            }
            if(fis != null){
                fis.close();
            }
        }
    }

    private void copyFileJar(String copyFile, String newFile){
        StringBuilder file = readFileJar(copyFile);
        writeTextToFile(newFile, file);
    }

    /**
     * 正常的赋值方法打成jar包就不行了，jar包需要用这个方法
     * @param uri 文件在jar包中的路径， / 开头
     * @return 返回文件内容
     */
    StringBuilder readFileJar(String uri){
        InputStream is = null;
        BufferedReader br = null;
        try {
            uri = uri.replace("\\", "/");
            is = IoUtil.class.getClass().getResourceAsStream(uri);
            br = new BufferedReader(new InputStreamReader(is));
            StringBuilder file = new StringBuilder();
            String s = "";
            while((s = br.readLine())!=null){
                file.append(s);
            }
            return file;
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }finally {
            if(br != null){
                try {
                    br.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if(is != null){
                try {
                    is.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * 拷贝图片
     * @param tempPath 文件临时路径
     * @param imgUri 图片路径
     */
    static int newImgAdd(String tempPath, String imgUri){
        String s = File.separator;
        File media = new File(tempPath + s + "word" + s + "media");
        media.mkdirs();
        int num = media.listFiles().length + 1;
        try {
            copyFile(imgUri, tempPath + s + "word" + s + "media" + s + "image" + num + "." + StrUtil.getImgFormat(imgUri));
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return num;
    }

    /**
     * 把图片添加到ContextType配置文件中
     * @param tempPath 文件夹临时路径
     * @param imgUri 图片地址
     */
    static void addContentTypes(String tempPath, String imgUri){
        StringBuilder types = readFile(tempPath + File.separator + Docx.DOCX_CONTENT);
        String imgFormat = StrUtil.getImgFormat(imgUri);
        if(types.indexOf(Docx.DOCX_CONTENT_TYPE_LEFT + imgFormat + Docx.DOCX_CONTENT_TYPE_RIGHT) < 0){
            StrUtil.replaceOnce(types, Docx.DOCX_CONTENT_PIC_GET, Docx.DOCX_CONTENT_PIC_GET + Docx.DOCX_CONTENT_TYPE_ADD_HEAD +
                    imgFormat + Docx.DOCX_CONTENT_TYPE_ADD_BODY + imgFormat + Docx.DOCX_CONTENT_TYPE_ADD_FOOT);
        }
        writeTextToFile(tempPath + File.separator + Docx.DOCX_CONTENT, types);
    }

    /**
     * 向文件资源关系文件中声明图片插入信息
     * @param tempPath 文件临时路径
     * @param imgUri 图片地址
     * @param imgId 图片资源id
     */
    static int addRelationship(String tempPath, String imgUri, int imgId){
        //获取已有资源数
        StringBuilder resources = readFile(tempPath + Docx.DOCX_RELATIONSHIP);
        int relNum = StrUtil.getRelationCount(resources.toString());
        //加上这个资源信息写回文件
        String imgFormat = StrUtil.getImgFormat(imgUri);
        String relationShip = Docx.DOCX_RELATIONSHIP_TEXT_1 + (relNum + 1) + Docx.DOCX_RELATIONSHIP_TEXT_2 + imgId +
                Docx.DOCX_RELATIONSHIP_TEXT_3 + imgFormat + Docx.DOCX_RELATIONSHIP_TEXT_4;
        StrUtil.replaceOnce(resources, Docx.DOCX_RELATIONSHIP_GET, relationShip + Docx.DOCX_RELATIONSHIP_GET);
        writeTextToFile(tempPath + Docx.DOCX_RELATIONSHIP, resources);
        return relNum + 1;
    }

    /**
     * 读取一个文件的内容
     * @param uri 文件路径
     * @return 文件内容
     */
    private static StringBuilder readFile(String uri){
        StringBuilder result = new StringBuilder();
        BufferedReader br = null;
        try{
            br = new BufferedReader(new FileReader(uri));
            String s = null;
            while((s = br.readLine()) != null){
                result.append(s);
            }
            br.close();
        }catch(Exception e){
            e.printStackTrace();
        }finally {
            if(br != null){
                try {
                    br.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return result;
    }

    static void zipMultiFile(String filepath ,String zippath) {
        ZipOutputStream zipOut = null;
        try {
            // 要被压缩的文件夹
            File file = new File(filepath);
            File zipFile = new File(zippath);
            zipOut = new ZipOutputStream(new FileOutputStream(zipFile));
            if(file.isDirectory()){
                File[] files = file.listFiles();
                if(files != null){
                    for(File fileSec : files){
                        recursionZip(zipOut, fileSec, "");
                    }
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            if(zipOut != null) {
                try {
                    zipOut.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * zip打包
     * @param zipOut 输出目录
     * @param file 文件路径
     * @param baseDir 是否压缩第一个文件夹作
     * @throws Exception 压缩出错
     */
    private static void recursionZip(ZipOutputStream zipOut, File file, String baseDir) throws Exception{
        if(file.isDirectory()){
            File[] files = file.listFiles();
            if(files == null){
                return;
            }
            for(File fileSec:files){
                recursionZip(zipOut, fileSec, baseDir + file.getName() + File.separator);
            }
        }else{
            byte[] buf = new byte[1024];
            InputStream input = null;
            try {
                input = new FileInputStream(file);
                zipOut.putNextEntry(new ZipEntry(baseDir + file.getName()));
                int len;
                while ((len = input.read(buf)) != -1) {
                    zipOut.write(buf, 0, len);
                }
            }catch (Exception e){
                e.printStackTrace();
            }finally {
                if(input != null){
                    input.close();
                }
            }
        }
    }

    /**
     * 把一段文本写入到文件中
     * @param fileUri 文件的地址
     * @param sb 需要写入的字符串
     */
    static void writeTextToFile(String fileUri, StringBuilder sb){
        FileOutputStream out = null;
        File file = new File(fileUri);
        try {
            out = new FileOutputStream(file);
            if(!file.exists()){
                file.createNewFile();
            }
            //注意编码，必须 UTF-8
            byte[] bytes = sb.toString().getBytes("utf-8");
            out.write(bytes);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            if(out != null){
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * 删除整个目录，包括子文件
     * @param file 目录地址
     */
    public static void deleteFile(File file){
        //判断文件不为null或文件目录存在
        if (file == null || !file.exists()){
            return;
        }
        //取得这个目录下的所有子文件对象
        File[] files = file.listFiles();
        if(files == null){
            return;
        }
        //遍历该目录下的文件对象
        for (File f: files){
            //打印文件名
            String name = file.getName();
            //判断子目录是否存在子目录,如果是文件则删除
            if (f.isDirectory()){
                deleteFile(f);
            }else {
                f.delete();
            }
        }
        //删除空文件夹  for循环已经把上一层节点的目录清空。
        file.delete();
    }

    /**
     * 获取图片宽度
     * @param file  图片文件
     * @return 宽度
     */
    static int getImgWidth(File file) {
        InputStream is = null;
        BufferedImage src;
        int ret = -1;
        try {
            is = new FileInputStream(file);
            src = javax.imageio.ImageIO.read(is);
            ret = src.getWidth(null);
            is.close();
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            if(is != null){
                try {
                    is.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return ret;
    }


    /**
     * 获取图片高度
     * @param file  图片文件
     * @return 高度
     */
    static int getImgHeight(File file) {
        InputStream is = null;
        BufferedImage src;
        int ret = -1;
        try {
            is = new FileInputStream(file);
            src = javax.imageio.ImageIO.read(is);
            ret = src.getHeight(null);
            is.close();
        } catch (Exception e) {
            e.printStackTrace();
        }finally {
            if(is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        return ret;
    }

    /**
     * 添加一段表格模板信息
     * @param template 表格模板id
     * @param tempPath 临时文件储存路径
     */
    void addTemplate(String template, String tempPath) {
        StringBuilder style = readFile(tempPath + Docx.DOCX_STYLE);
        if(style.indexOf(Docx.DOCX_STYLE_HEAD + template) < 0){
            //此模板不在文件中
            StrUtil.replaceOnce(style, Docx.DOCX_STYLE_ADD, readFileJar(File.separator + "style" + File.separator + template + ".tbl").toString() + Docx.DOCX_STYLE_ADD);
            writeTextToFile(tempPath + Docx.DOCX_STYLE, style);
        }
    }

    /**
     * 在content-type文件中添加页眉页脚声明
     * @param tempPath 文件临时路径
     */
    void contentAddPaperOut(String tempPath){
        StringBuilder types = readFile(tempPath + File.separator + Docx.DOCX_CONTENT);
        StrUtil.replaceOnce(types, "<!--paperOut-->", readFileJar(File.separator + "content.ppo").toString());
        writeTextToFile(tempPath + File.separator + Docx.DOCX_CONTENT, types);
    }

    /**
     * 添加页眉页脚资源
     * @param tempPath 临时路径
     * @return 下一个资源的序号
     */
    static int addPaperOutRelationship(String tempPath){
        //获取已有资源数
        StringBuilder resources = readFile(tempPath + Docx.DOCX_RELATIONSHIP);
        int relNum = StrUtil.getRelationCount(resources.toString());
        //拼接数据
        String xml = "<Relationship Id=\"rId" +
                (relNum + 1) +
                "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/header\" Target=\"header1.xml\"/>" +
                "<Relationship Id=\"rId" +
                (relNum + 2) +
                "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer\" Target=\"footer1.xml\"/>" +
                "<Relationship Id=\"rId" +
                (relNum + 3) +
                "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes\" Target=\"endnotes.xml\"/>" +
                "<Relationship Id=\"rId" +
                (relNum + 4) +
                "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes\" Target=\"footnotes.xml\"/>";
        StrUtil.replace(resources, "<!--paperOut-->", xml);
        //写出数据
        writeTextToFile(tempPath + Docx.DOCX_RELATIONSHIP, resources);
        return relNum + 1;
    }

    /**
     * 在样式文件中添加页眉页脚信息
     * @param styleXml style的xml信息
     * @param tempPath 临时文件目录
     */
    static void paperOutAddStyle(StringBuilder styleXml, String tempPath) {
        StringBuilder style = readFile(tempPath + Docx.DOCX_STYLE);
        StrUtil.replaceOnce(style, Docx.DOCX_STYLE_ADD, styleXml.toString() + Docx.DOCX_STYLE_ADD);
        writeTextToFile(tempPath + Docx.DOCX_STYLE, style);
    }

    void paperOutAddFile(String tempPath, String uri){
        String s = File.separator;
        try {
            copyFileJar( s + "word" + s + "footnotes.xml2", tempPath + s + "word" + s + "footnotes.xml");
            copyFileJar(s + "word" + s + "endnotes.xml2", tempPath + s + "word" + s + "endnotes.xml");
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    void copyPaperOutSetting(String tempPath) {
        StringBuilder set = readFileJar(File.separator + "word" + File.separator + "settings.ppo");
        writeTextToFile(tempPath + File.separator + "word" + File.separator + "settings.xml", set);
    }

    static void contentCreate(String tempPath) {
        StringBuilder content = readFile(tempPath + File.separator + Docx.DOCX_CONTENT);
        StrUtil.replace(content, "<!--picGet-->", "");
        StrUtil.replace(content, "\"?><Types", "\"?>\n<Types");
        writeTextToFile(tempPath + File.separator + Docx.DOCX_CONTENT, content);
    }

    static void wordResCreate(String tempPath) {
        StringBuilder word = readFile(tempPath + Docx.DOCX_RELATIONSHIP);
        StrUtil.replace(word, "<!--res-->", "");
        writeTextToFile(tempPath + Docx.DOCX_RELATIONSHIP, word);
    }

    static void wordStyleCreate(String tempPath) {
        StringBuilder word = readFile(tempPath + Docx.DOCX_STYLE);
        StrUtil.replace(word, "<!--add-->", "");
        writeTextToFile(tempPath + Docx.DOCX_STYLE, word);
    }
}
