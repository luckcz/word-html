package com.mayikt.wordhtml.html2word;

import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

/**
 * @author ChenZhuang
 * @ClassName Html2Word3
 * @description TODO
 * @Date 2019/9/2 10:18
 * @Version 1.0
 *
 *  这个html转换word是成功的，图片也有
 */
public class Html2Word3 {
    public static void main(String[] args) throws IOException {
        //html拼接出word内容
        String content = "<html>";
        //String title = "标题";
        String cx = getBodyString();
        //html拼接出word内容  这个标题没必要要
        /*content+="<div style=\"text-align: center\"><span style=\"font-size: 24px\"><span style=\"font-family: 黑体\">"+title+"<br /> <br /> </span></span></div>";
        content+="<div style=\"text-align: left\"><span >"+cx+"<br /> <br /> </span></span></div>";*/
        //插入分页符
        content += "<span lang=EN-US style='font-size:12.0pt;line-height:150%;mso-fareast-font-family:宋体;mso-font-kerning:1.0pt;mso-ansi-language:EN-US;mso-fareast-language:ZH-CN;mso-bidi-language:AR-SA'><br clear=all style='page-break-before:always'></span>";
        content += "<p class=MsoNormal style='line-height:150%'><span lang=EN-US style='font-size:12.0pt;line-height:150%'><o:p> </o:p></span></p>";
        content += "</html>";
        System.out.println(content);
        byte b[] = content.getBytes();
        ByteArrayInputStream bais = new ByteArrayInputStream(b);
        POIFSFileSystem poifs = new POIFSFileSystem();
        DirectoryEntry directory = poifs.getRoot();
        DocumentEntry documentEntry = directory.createDocument("WordDocument", bais);
        //输出文件,
        /*一般这个是会传到前端去直接下载的，所以这个用的比较多
        response.reset();
        response.setHeader("Content-Disposition",
                "attachment;filename=" +
                        new String( (name + ".doc").getBytes(),
                                "iso-8859-1"));
        response.setContentType("application/msword");
        OutputStream ostream = response.getOutputStream();*/
        //输出文件的话，new一个文件流
        FileOutputStream ostream = new FileOutputStream("D:/cccc.doc");
        poifs.writeFilesystem(ostream);
        ostream.flush();
        ostream.close();
        bais.close();
    }

    public static String getBodyString() throws IOException {
        File file = new File("D:\\cz\\壮\\桌面资料\\工作资料\\流程复盘节点确定.html");
        InputStream in = new FileInputStream(file);
        InputStreamReader reader = new InputStreamReader(in);
        BufferedReader br = new BufferedReader(reader);
        String s = null;
        StringBuffer buff = new StringBuffer();
        while ((s = br.readLine()) != null) {
            buff.append(s);
        }
        br.close();
        reader.close();
        in.close();
        System.out.println(buff.toString());
        buff.delete(0,buff.indexOf("<head>"));
        buff.deleteCharAt(buff.indexOf("</html>"));
        /*StringBuffer buffStyle = new StringBuffer();
        //截取样式代码
        buffStyle.append(buff.substring(buff.indexOf("<style type=\"text/css\">") + 23, buff.indexOf("style", buff.indexOf("<style type=\"text/css\">") + 23) - 2));
        System.out.println(buffStyle);
        //截取body代码
        String body = buff.substring(buff.indexOf("<body"), buff.indexOf("</body") + 7);
        body = body.replaceAll("body", "div");
        StringBuffer bodyBuffer = new StringBuffer(body);
        System.out.println(bodyBuffer);
        String[] split = buffStyle.toString().split("}");
        Map<String, String> styleMap = new HashMap<>();
        for (String s1 : split) {
            System.out.println(s1);
            String[] split1 = s1.split("\\{");
            styleMap.put(split1[0].substring(1), split1[1]);
        }
        Set<String> strings = styleMap.keySet();
        for (String key : strings) {
            System.out.print("key : " + key);
            System.out.println("   value : " + styleMap.get(key));
            //将嵌入样式转换为行内样式
            if (bodyBuffer.toString().contains(key)) {
                int length = bodyBuffer.toString().split(key).length - 1;
                int temp = 0;
                for (int i = 0; i < length; i++) {
                    temp = bodyBuffer.indexOf(key, temp);
                    //这个是每次查询到的位置，判断此标签中是否添加了style标签
                    String isContaionStyle = bodyBuffer.substring(temp, bodyBuffer.indexOf(">", temp));
                    if (isContaionStyle.contains("style")) {
                        //代表已经存在此style，那么直接加进去就好了
                        //首先找到style的位置
                        int styleTemp = bodyBuffer.indexOf("style", temp);
                        bodyBuffer.insert(styleTemp + 7, styleMap.get(key));
                    } else {
                        //代表没有style，那么直接插入style
                        int styleIndex = bodyBuffer.indexOf("\"", temp);
                        bodyBuffer.insert(styleIndex + 1, " style=\"" + styleMap.get(key) + "\"");
                    }
                    temp++;
                }
            }
            System.out.println(bodyBuffer.toString());
        }*/
        System.out.println(buff.toString());
        return buff.toString();
    }
}