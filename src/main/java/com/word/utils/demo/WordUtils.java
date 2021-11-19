package com.word.utils.demo;

import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

/*
 * @Author cyf
 * @Description
 * @Date 2021/11/19
 **/
public class WordUtils {


    public static void main(String[] args) {

        String inputSrc = "C:\\Users\\Administrator.DESKTOP-EFOTAD8\\Desktop\\test.docx";
        String outputSrc =  "C:\\Users\\Administrator.DESKTOP-EFOTAD8\\Desktop\\test02.docx";
        Map<String, String> contentMaps = new HashMap<>();
        contentMaps.put("${name}","chen");
        contentMaps.put("&img","C:\\Users\\Administrator.DESKTOP-EFOTAD8\\Desktop\\img.jpg");
        copyWord(inputSrc,outputSrc,contentMaps);

    }

    private static boolean copyWord(String inputSrc, String outputSrc, Map<String, String> contentMaps) {

        boolean success = true;
        try {
            XWPFDocument document = new XWPFDocument(POIXMLDocument.openPackage(inputSrc));
            copyText(document,contentMaps);
            FileOutputStream outputStream = new FileOutputStream(outputSrc);
            document.write(outputStream);
            outputStream.close();

        } catch (IOException e) {
            e.printStackTrace();
            success = false;
        }

        return success;

    }

    private static void copyText(XWPFDocument document, Map<String, String> contentMaps) {
        //获取 word 所有的段落
        List<XWPFParagraph> paragraphList = document.getParagraphs();
        for (XWPFParagraph paragraph : paragraphList) {
            List<XWPFRun> runs = paragraph.getRuns();
            for (XWPFRun run : runs) {
                // run 一个单词
                String str = run.toString();
                if (checkText(str)){
                    //包含 $则进行替换
                    //pos:位置，如果不要这个参数则是追加字符串
                    run.setText(contentMaps.get(str),0);
                }else if (checkImg(str)){
                    setImg(run,paragraph,contentMaps);
                }

            }
        }
    }

    private static boolean checkImg(String str) {
        if (str.contains("&")){
            return true;
        }
        return false;
    }

    public static void setImg(XWPFRun run,XWPFParagraph paragraph,Map<String, String> contentMaps){
        String str = run.toString();
        try {
            if ("&img".equals(str)){
                String imgSrc = contentMaps.get(str);
                FileInputStream is = new FileInputStream(imgSrc);
                run.setBold(true);
                paragraph.setAlignment(ParagraphAlignment.CENTER);
                run.addBreak();
                run.addPicture(is,XWPFDocument.PICTURE_TYPE_JPEG,imgSrc, Units.toEMU(200),Units.toEMU(200));
                is.close();
            }

        }catch (Exception e){
            e.printStackTrace();
        }
    }


    //判断 run 有没有包含 $
    private static boolean checkText(String str) {
        if (str.contains("$")){
            return true;
        }
        return false;
    }

}
