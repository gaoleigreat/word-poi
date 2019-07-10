package com.legao.word.poi;

import com.legao.word.poi.vo.WordPictureVO;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

public class Test {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        //模板文件
        InputStream fileInputStream = new FileInputStream("C:/Users/xiaodao/Desktop/xxx_new.docx");
        XWPFDocument document = new XWPFDocument(fileInputStream);
        /**
         * 存放更换变量的容器
         * key为模板里需要替换的值
         * value为 key替换后的值
         */

        Map<String, String> params = new HashMap<>();
        /**
         * 存放更换图片的容器
         * key为模板里需要替换的值
         * value为 key替换后的图片
         */
        Map<String, WordPictureVO> pictures = new HashMap<>();


        /**需要添加的图片封装对象
         * 参数分别为:
         * 图片输入流
         * 图片类型
         * 图片名
         * 图片宽度和高度
         */
        WordPictureVO wordPictureVO = new WordPictureVO(new FileInputStream("C:/Users/xiaodao/Desktop/1.png"), XwpfTUtil.PICTURE_TYPE_PNG, "", 200, 100);
        WordPictureVO wordPictureVO1 = new WordPictureVO(new FileInputStream("C:/Users/xiaodao/Desktop/1.png"), 6, "", 200, 100);

        params.put("${name}", "高磊");
        params.put("${address}", "陕西省乾县");
        pictures.put("${picture}", wordPictureVO);
        pictures.put("${picture1}", wordPictureVO1);

        XwpfTUtil xwpfTUtil = new XwpfTUtil();
        //调用方法进行生成文档
        xwpfTUtil.repaceWordAndExcel(document, params, pictures);


        //关闭输出流
        OutputStream os = new FileOutputStream("C:/Users/xiaodao/Desktop/xxx_new2.docx");
        document.write(os);
        //关闭输入流
        fileInputStream.close();
        os.close();


    }
}
