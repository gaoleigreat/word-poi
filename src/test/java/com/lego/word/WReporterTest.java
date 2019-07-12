package com.lego.word;

import com.lego.word.element.WObject;
import com.lego.word.element.WPic;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Set;

public class WReporterTest {
    private WReporter wReporter = new WReporter("C:/Users/xiaodao/Desktop/xxx_new.docx");

    WReporterTest() throws IOException {
    }

    /**
     * 在表格中插入数据
     * @throws Exception
     */
    @Test
    void export() throws Exception {

        List<WObject> uSers = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            MyTable uSer = new MyTable();
            uSer.setAge(i);
            uSer.setName("xiaodao" + i);
            uSer.setSex("nana" + i);
            WPic wPic = new WPic();
            wPic.setHeight(100);
            wPic.setWidth(180);
            wPic.setType("png");
            //可以图片url也可以是图片路径
            wPic.setUrl("https://timgsa.baidu.com/timg?image&quality=80&size=b9999_10000&sec=1562912750192&di=8e6e37f2fcc76f049c72bb0d1b9d10f8&imgtype=0&src=http%3A%2F%2Fb-ssl.duitang.com%2Fuploads%2Fitem%2F201503%2F07%2F20150307213403_rQrCt.thumb.700_0.jpeg");
            uSer.setBir(wPic);
            uSers.add(uSer);
        }
        //第二个参数表示第几张表，从0开始，第三个参数表示模板行是第几行，从0开始为第一行
        wReporter.export(uSers,1,1);
        wReporter.generate("C:/Users/xiaodao/Desktop/result.docx");
    }

    /**
     * 替换文档里面的内容
     */
    @Test
    void export1() throws Exception {
        MyDoc myDoc = new MyDoc();
        myDoc.setName("gaolei");
        myDoc.setAddress("陕西");
        WPic wPic = new WPic();
        wPic.setHeight(100);
        wPic.setWidth(180);
        wPic.setType("png");
        wPic.setUrl("https://timgsa.baidu.com/timg?image&quality=80&size=b9999_10000&sec=1562912750192&di=8e6e37f2fcc76f049c72bb0d1b9d10f8&imgtype=0&src=http%3A%2F%2Fb-ssl.duitang.com%2Fuploads%2Fitem%2F201503%2F07%2F20150307213403_rQrCt.thumb.700_0.jpeg");
        myDoc.setPicture(wPic);
        wReporter.export(myDoc);
        wReporter.generate("C:/Users/xiaodao/Desktop/result.docx");
    }

    /**
     * 替换表格中数据
     * @throws Exception
     */
    @Test
    void export2() throws Exception {

        MyDoc myDoc = new MyDoc();
        myDoc.setName("gaolei");
        myDoc.setAddress("陕西");
        WPic wPic = new WPic();
        wPic.setHeight(100);
        wPic.setWidth(180);
        wPic.setType("png");
        wPic.setUrl("https://timgsa.baidu.com/timg?image&quality=80&size=b9999_10000&sec=1562912750192&di=8e6e37f2fcc76f049c72bb0d1b9d10f8&imgtype=0&src=http%3A%2F%2Fb-ssl.duitang.com%2Fuploads%2Fitem%2F201503%2F07%2F20150307213403_rQrCt.thumb.700_0.jpeg");
        myDoc.setPicture(wPic);
        wReporter.export(myDoc,0);
        wReporter.generate("C:/Users/xiaodao/Desktop/result.docx");
    }

    /**
     * 查看文档中替换的参数
     * @throws Exception
     */
    @Test
    void findInPara() throws Exception {
        Set<String> set = wReporter.findInPara();
        System.out.println("");
    }

    /**
     * 查看表格中替换的参数
     * @throws Exception
     * tableindex表示第几张表，从0开始 -1表示所有表
     */
    @Test
    void findInTable() throws Exception {
        Set<String> set = wReporter.findInTable(0);
        System.out.println("");
    }

}