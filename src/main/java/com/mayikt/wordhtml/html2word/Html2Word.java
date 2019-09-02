package com.mayikt.wordhtml.html2word;

/**
 * @author ChenZhuang
 * @ClassName Html2Word
 * @description TODO
 * @Date 2019/8/29 9:51
 * @Version 1.0
 */
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;



public class Html2Word {

    public static boolean writeWordFile() {

        boolean w = false;
        String path = "D:/";

        try {
            if (!"".equals(path)) {

                // 检查目录是否存在
                File fileDir = new File(path);
                if (fileDir.exists()) {

                    // 生成临时文件名称
                    String fileName = "a.doc";
                    String content = "<!DOCTYPE html>\n" +
                            "<html lang=\"en\">\n" +
                            "<head>\n" +
                            "    <meta charset=\"UTF-8\">\n" +
                            "    <title>Title</title>\n" +
                            "</head>\n" +
                            "<body>\n" +
                            "<div class=\"b1 b2\" style=\"white-space-collapsing:preserve;margin: 1.0in 1.25in 1.0in 1.25in;\">\n" +
                            "<p class=\"p1\" style=\"margin-top:0.2361111in;margin-bottom:0.22916667in;text-align:center;hyphenate:auto;keep-together.within-page:always;keep-with-next.within-page:always;font-family:Calibri;font-size:22pt;\">\n" +
                            "    <span class=\"s1\" style=\"font-weight:bold;\">一级标题</span></p>\n" +
                            "    <p class=\"p2\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:14pt;\">\n" +
                            "\t<span class=\"s1\" style=\"font-weight:bold;\">流程节点分类</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>唯一人审核。</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>部门领导审核。</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>公司人员审核。</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>公司上一层人员审核。</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>注：节点提交时，当流程节点未配置人员，提示为&ldquo;配置人员，请配置审核人员后提交&rdquo;。</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>配好人员之后，流程可以继续往下走。</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p2\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:14pt;\">\n" +
                            "\t<span class=\"s1\" style=\"font-weight:bold;\">流程梳理</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p4\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:11pt;\">\n" +
                            "\t<span class=\"s1\" style=\"font-weight:bold;\">备案流程</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>厂长审核------公司人员审核。</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>事业部设备副总审核-----公司人员审核。注：往上找一层或多层。</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>工厂设备科审核-------公司人员审核</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>资产部安环科审核------唯一人审核</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>集团技改办填写预算-----唯一人审核</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>总经理审核-------公司人员审核。注：往上找一层或多层。</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>内勤报董事长审核------唯一人审核</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p4\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:11pt;\"><span class=\"s1\"\n" +
                            "                                                                                                      style=\"font-weight:bold;\">采购流程</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>集团技改办确认采购金额---------唯一人审核</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>备案项目技术负责人填写技术意见---------a、备案单技术负责人审核；b、</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>厂长审核-------公司人员审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>设备副总审核------公司人员审核。注：往上找一层或多层</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>资产部审核-------唯一人审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>内勤报董事长审核-------唯一人审核</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p4\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:11pt;\"><span class=\"s1\"\n" +
                            "                                                                                                      style=\"font-weight:bold;\">设备闲置申请单</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span class=\"s2\"\n" +
                            "                                                                                                      style=\"color:blue;\">使用部门负责人审核------部门领导审核-----表单使用部门负责人审核</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>设备科审核-------公司人员审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>厂长审核-------公司人员审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>设备副总审核------公司人员审核。注：往上找一层或多层</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>资产部审核------唯一人审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p4\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:11pt;\"><span class=\"s1\"\n" +
                            "                                                                                                      style=\"font-weight:bold;\">闲置设备启动申请单</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>资产部审核-------唯一人审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p4\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:11pt;\"><span class=\"s1\"\n" +
                            "                                                                                                      style=\"font-weight:bold;\">闲置设备处置申请单</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>设备科负责人审核------公司人员审核</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>工厂厂长审核-----公司人员审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>分管设备副总审核------公司人员审核。注：往上找一层或多层</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>事业部总经理审核-----公司人员审核。注：往上找一层或多层</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>财务负责人审核------唯一人审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>资产部审核------唯一人审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>内勤报董事长审核------唯一人审核</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p4\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:11pt;\"><span class=\"s1\"\n" +
                            "                                                                                                      style=\"font-weight:bold;\">电机处置评审单</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>设备科负责人审核-----公司人员审核</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>工厂厂长审核-----公司人员审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>资产部审核------唯一人审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p4\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:11pt;\"><span class=\"s1\"\n" +
                            "                                                                                                      style=\"font-weight:bold;\">设备出厂维修单</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span class=\"s2\"\n" +
                            "                                                                                                      style=\"color:blue;\">经办部门/部门领导签字-------部门领导审核----发起人部门领导</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>调出厂签字-----公司人员审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>设备副总审核-----公司人员审核。注：往上找一层或多层</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>资产部副部长审核----唯一人审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p4\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:11pt;\"><span class=\"s1\"\n" +
                            "                                                                                                      style=\"font-weight:bold;\">设备厂内迁移单</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span class=\"s2\"\n" +
                            "                                                                                                      style=\"color:blue;\">经办部门/部门领导签字-----部门领导审核--发起人部门领导</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>设备副总审核------公司人员审核。注：往上找一层或多层</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>资产部副部长审核-----唯一人审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><img\n" +
                            "            src=\"D:\\cz\\%E6%A1%8C%E9%9D%A2%E8%B5%84%E6%96%99\\%E5%B7%A5%E4%BD%9C%E8%B5%84%E6%96%99\\/image/0.png\"\n" +
                            "            style=\"width:5.7541666in;height:2.195139in;vertical-align:text-bottom;\"></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p4\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:11pt;\"><span class=\"s1\"\n" +
                            "                                                                                                      style=\"font-weight:bold;\">设备调拨单</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>经办部门/部门领导签字------部门领导审核-----发起人部门领导</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>调出厂签字-----公司人员审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>设备副总签字-----公司人员审核。注：往上找一层或多层</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>资产部副部长审核------唯一人审核</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span class=\"s2\"\n" +
                            "                                                                                                      style=\"color:blue;\">调入厂设备科科长审核------表单上调入工厂的公司人员审核</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p4\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:11pt;\"><span class=\"s1\"\n" +
                            "                                                                                                      style=\"font-weight:bold;\">调拨质量信息报告</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span class=\"s2\"\n" +
                            "                                                                                                      style=\"color:blue;\">车间主任-----调入部门的部门领导审核。---表单使用部门</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>设备科长----公司人员审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>工厂厂长-----公司人员审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p4\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:11pt;\"><span class=\"s1\"\n" +
                            "                                                                                                      style=\"font-weight:bold;\">新增配件申请单</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>设备科审核------公司人员审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>资产部审核-----唯一人审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p4\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:11pt;\"><span class=\"s1\"\n" +
                            "                                                                                                      style=\"font-weight:bold;\">设备报修单</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>厂长审核------公司人员审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><span>设备副总审核-------公司人员审核。注：往上找一层或多层</span>\n" +
                            "    </p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\">\n" +
                            "        <span>资产部审核------唯一人审核</span></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"><img\n" +
                            "            src=\"D:\\cz\\%E6%A1%8C%E9%9D%A2%E8%B5%84%E6%96%99\\%E5%B7%A5%E4%BD%9C%E8%B5%84%E6%96%99\\/image/71870.png\"\n" +
                            "            style=\"width:5.7541666in;height:2.195139in;vertical-align:text-bottom;\"></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p>\n" +
                            "    <p class=\"p3\" style=\"text-align:justify;hyphenate:auto;font-family:Calibri;font-size:10pt;\"></p></div>\n" +
                            "</body>\n" +
                            "</html>";

                    byte b[] = content.getBytes();
                    ByteArrayInputStream bais = new ByteArrayInputStream(b);
                    POIFSFileSystem poifs = new POIFSFileSystem();
                    DirectoryEntry directory = poifs.getRoot();
                    DocumentEntry documentEntry = directory.createDocument("WordDocument", bais);
                    FileOutputStream ostream = new FileOutputStream(path+ fileName);
                    poifs.writeFilesystem(ostream);
                    bais.close();
                    ostream.close();

                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        return w;
    }

    public static void main(String[] args){
        writeWordFile();
    }

}