package com.mayikt.wordhtml.word2html2;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.apache.poi.xwpf.converter.core.FileImageExtractor;
import org.apache.poi.xwpf.converter.core.FileURIResolver;
import org.apache.poi.xwpf.converter.xhtml.XHTMLConverter;
import org.apache.poi.xwpf.converter.xhtml.XHTMLOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.Test;
import org.w3c.dom.Document;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.*;
import java.util.*;

/**
 * @author ChenZhuang
 * @ClassName Word2Html
 * @description TODO
 * @Date 2019/8/27 11:10
 * @Version 1.0
 */
public class Word2Html {
    /**
     * 将Word2007+转成Html
     *
     * @throws Exception
     */
    @Test
    public void word2007ToHtml() throws Exception {
        String filePath = "E:/学习、练习数据文件夹/test/";
        String fileName = "SpringIOC解析.docx";
        String htmlName = "SpringIOC解析.html";
        final String file = filePath + fileName;
        File f = new File(file);
        if (!f.exists()) {
            System.out.println("Sorry File does not Exists!");
        } else {
            /* 判断是否为docx文件 */
            if (f.getName().endsWith(".docx") || f.getName().endsWith(".DOCX")) {
                // 1)加载word文档生成XWPFDocument对象
                FileInputStream in = new FileInputStream(f);
                XWPFDocument document = new XWPFDocument(in);
                // 2)解析XHTML配置（这里设置IURIResolver来设置图片存放的目录）
                File imageFolderFile = new File(filePath);
                XHTMLOptions options = XHTMLOptions.create().URIResolver(new FileURIResolver(imageFolderFile));
                options.setExtractor(new FileImageExtractor(imageFolderFile));
                options.setIgnoreStylesIfUnused(false);
                options.setFragment(true);
                // 3)将XWPFDocument转换成XHTML
                FileOutputStream out = new FileOutputStream(new File(filePath + htmlName));
                XHTMLConverter.getInstance().convert(document, out, options);
                //也可以使用字符数组流获取解析的内容
                //ByteArrayOutputStream baos = new ByteArrayOutputStream();
                //XHTMLConverter.getInstance().convert(document, baos, options);
                //String content = baos.toString();
                //System.out.println(content);
                //baos.close();
            } else {
                System.out.println("Enter only as MS Office 2007+ files");
            }
        }
    }

    /**
     * word2003-2007转换成html
     * @throws Exception
     */
    @Test
    public void wordToHtml() throws Exception {
        String filePath = "D:\\cz\\桌面资料\\工作资料\\";
        String fileName = "流程复盘节点确定.doc";
        String htmlName = "流程复盘节点确定.html";
        final String imagePath = filePath + "/image/";
        final String file = filePath + fileName;
        InputStream input = new FileInputStream(new File(file));
        HWPFDocument wordDocument = new HWPFDocument(input);
        WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
                DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument());
        //设置图片存储位置
        wordToHtmlConverter.setPicturesManager(new PicturesManager() {

            public String savePicture(byte[] content, PictureType pictureType, String suggestedName, float widthInches,
                                      float heightInches) {
                File imgPath=new File(imagePath);
                if (!imgPath.exists()) {//目录不存在则创建目录
                    imgPath.mkdirs();
                }
                File file = new File(imagePath+suggestedName);
                try {
                    FileOutputStream os = null;
                    try {
                        os = new FileOutputStream(file);
                    } catch (FileNotFoundException e) {
                        e.printStackTrace();
                    }
                    os.write(content);
                    os.close();
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                }
                return imagePath+suggestedName;
            }
        });

        //解析word文档
        wordToHtmlConverter.processDocument(wordDocument);
        Document htmlDocument = wordToHtmlConverter.getDocument();
        File htmlFile = new File(filePath+htmlName);
        FileOutputStream outStream = new FileOutputStream(htmlFile);
        //也可以使用字符数组流获取解析的内容
        //ByteArrayOutputStream baos = new ByteArrayOutputStream();
        //OutputStream outStream = new BufferedOutputStream(baos);
        DOMSource domSource = new DOMSource(htmlDocument);
        StreamResult streamResult = new StreamResult(outStream);
        TransformerFactory factory = TransformerFactory.newInstance();
        Transformer serializer = factory.newTransformer();
        serializer.setOutputProperty(OutputKeys.ENCODING,"utf-8");
        serializer.setOutputProperty(OutputKeys.INDENT, "yes");
        serializer.setOutputProperty(OutputKeys.METHOD, "html");
        serializer.transform(domSource, streamResult);
        //也可以使用字符数组流获取解析的内容
        //String content = baos.toString();
        //System.out.println(content);
        //baos.close();
        outStream.close();
        System.out.println("转换完成");
    }

    //读取html文件,读取之后内嵌式css转内联式css
    @Test
    public void fun() throws IOException {
        File file = new File("D:\\cz\\桌面资料\\工作资料\\流程复盘节点确定.html");
        InputStream in = new FileInputStream(file);
        InputStreamReader reader = new InputStreamReader(in);
        BufferedReader br = new BufferedReader(reader);
        String s = null ;
        StringBuffer buff = new StringBuffer();
        while ((s = br.readLine()) != null){
            buff.append(s);
        }
        br.close();
        reader.close();
        in.close();
        System.out.println(buff.toString());
        StringBuffer buffStyle = new StringBuffer();
        //截取样式代码
        buffStyle.append(buff.substring(buff.indexOf("<style type=\"text/css\">") +23 ,buff.indexOf("style",buff.indexOf("<style type=\"text/css\">") +23 )-2));
        System.out.println(buffStyle);
        //截取body代码
        String body = buff.substring(buff.indexOf("<body"),buff.indexOf("</body")+7);
        body = body.replaceAll("body","div");
        StringBuffer bodyBuffer = new StringBuffer(body);
        System.out.println(bodyBuffer);
        String[] split = buffStyle.toString().split("}");
        Map<String,String> styleMap = new HashMap<>();
        for (String s1 : split) {
            System.out.println(s1);
            String[] split1 = s1.split("\\{");
            styleMap.put(split1[0].substring(1),split1[1]);
        }
        Set<String> strings = styleMap.keySet();
        for (String key : strings) {
            System.out.print("key : "+key);
            System.out.println("   value : "+styleMap.get(key));
            //将嵌入样式转换为行内样式
            if(bodyBuffer.toString().contains(key)){
                int length = bodyBuffer.toString().split(key).length - 1 ;
                int temp = 0 ;
                for (int i = 0 ; i < length ; i++){
                    temp = bodyBuffer.indexOf(key,temp);
                    //这个是每次查询到的位置，判断此标签中是否添加了style标签
                    String isContaionStyle = bodyBuffer.substring(temp,bodyBuffer.indexOf(">",temp));
                    if(isContaionStyle.contains("style")){
                        //代表已经存在此style，那么直接加进去就好了
                        //首先找到style的位置
                        int styleTemp = bodyBuffer.indexOf("style",temp);
                        bodyBuffer.insert(styleTemp+7,styleMap.get(key));
                    }else{
                        //代表没有style，那么直接插入style
                        int styleIndex = bodyBuffer.indexOf("\"",temp);
                        bodyBuffer.insert(styleIndex+1," style=\""+styleMap.get(key)+"\"");
                    }
                    temp++;
                }
            }
        }
        System.out.println(bodyBuffer.toString());
    }

    @Test
    public void fun2(){
        String str = "abab>abcbcbcbcbcdbd><div class=''></div>";
        StringBuffer buffer = new StringBuffer(str);
        int length = str.split("b").length - 1;
        System.out.println(length);
        int temp = 0 ;
        List<Integer> list = new ArrayList<>();
        for (int i = 0 ; i < length ; i++){
            temp = str.indexOf("b",temp);
            list.add(temp);
            temp++;
        }
        System.out.println(list);
        System.out.println(str.substring(6,str.indexOf(">",6)));
        String substring = buffer.toString().substring(1, 6);
        System.out.println(substring);
        System.out.println(buffer);
        buffer.insert(1,"9876");
        System.out.println(buffer);
    }
}
