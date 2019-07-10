package com.legao.word.poi;

import com.legao.word.poi.vo.WordPictureVO;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Slf4j
public class XwpfTUtil {

    public static final int PICTURE_TYPE_EMF = 2;

    /**
     * Windows Meta File
     */
    public static final int PICTURE_TYPE_WMF = 3;

    /**
     * Mac PICT format
     */
    public static final int PICTURE_TYPE_PICT = 4;

    /**
     * JPEG format
     */
    public static final int PICTURE_TYPE_JPEG = 5;

    /**
     * PNG format
     */
    public static final int PICTURE_TYPE_PNG = 6;

    /**
     * Device independent bitmap
     */
    public static final int PICTURE_TYPE_DIB = 7;

    /**
     * GIF image format
     */
    public static final int PICTURE_TYPE_GIF = 8;

    /**
     * Tag Image File (.tiff)
     */
    public static final int PICTURE_TYPE_TIFF = 9;

    /**
     * Encapsulated Postscript (.eps)
     */
    public static final int PICTURE_TYPE_EPS = 10;


    /**
     * Windows Bitmap (.bmp)
     */
    public static final int PICTURE_TYPE_BMP = 11;

    /**
     * WordPerfect graphics (.wpg)
     */
    public static final int PICTURE_TYPE_WPG = 12;


    /**
     * 替换段落里面的变量
     *
     * @param doc    要替换的文档
     * @param params 参数
     */
    public void replaceInPara(XWPFDocument doc, Map<String, String> params, Map<String, WordPictureVO> pictures, Map<String, String> picIds) throws IOException, InvalidFormatException {
        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        XWPFParagraph para;
        while (iterator.hasNext()) {
            para = iterator.next();
            this.replaceInPara(para, params, pictures, picIds);
        }
    }

    /**
     * 替换段落里面的变量
     *
     * @param para   要替换的段落
     * @param params 参数
     */
    public void replaceInPara(XWPFParagraph para, Map<String, String> params, Map<String, WordPictureVO> pictures, Map<String, String> picIds) throws IOException, InvalidFormatException {
        List<XWPFRun> runs;
        if (this.matcher(para.getParagraphText()).find()) {
            runs = para.getRuns();
            String str = "";
            for (int i = 0; i < runs.size(); i++) {
                XWPFRun run = runs.get(i);
                String runText = run.toString();
                str = str + runText;

            }

            while (para.getRuns().size() != 0) {
                para.removeRun(para.getRuns().size() - 1);
            }
            List<String> keys = getKeyListBySet(params.keySet());
            keys.addAll(getKeyListBySet(picIds.keySet()));
            List<String> runString = WordStringUtil.splitStringList(Arrays.asList(str), keys);
            for (String s : runString) {
                if (params != null && params.containsKey(s)) {
                    para.createRun().setText(params.get(s));
                } else if (pictures != null && pictures.containsKey(s)) {
                    addPictureToRun(para.createRun(), picIds.get(s), XWPFDocument.PICTURE_TYPE_JPEG, 100, 100);
                } else {
                    para.createRun().setText(s);
                }

            }

        }
    }


    /**
     * 替换表格里面的变量
     *
     * @param doc    要替换的文档
     * @param params 参数
     */
    public void replaceInTable(XWPFDocument doc, Map<String, String> params, Map<String, WordPictureVO> pictures, Map<String, String> picIds) throws IOException, InvalidFormatException {
        Iterator<XWPFTable> iterator = doc.getTablesIterator();
        XWPFTable table;
        List<XWPFTableRow> rows;
        List<XWPFTableCell> cells;
        List<XWPFParagraph> paras;
        while (iterator.hasNext()) {
            table = iterator.next();
            rows = table.getRows();
            for (XWPFTableRow row : rows) {
                cells = row.getTableCells();
                for (XWPFTableCell cell : cells) {
                    paras = cell.getParagraphs();
                    for (XWPFParagraph para : paras) {
                        this.replaceInPara(para, params, pictures, picIds);
                    }
                }
            }
        }
    }

    /**
     * 正则匹配字符串
     *
     * @param str
     * @return
     */
    private Matcher matcher(String str) {
        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }

    /**
     * 关闭输入流
     *
     * @param is
     */
    public void close(InputStream is) {
        if (is != null) {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 关闭输出流
     *
     * @param os
     */
    public void close(OutputStream os) {
        if (os != null) {
            try {
                os.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


    /**
     * 根据map获取map包含的key,返回set
     *
     * @param map
     * @return
     */
    public static Set<String> getKeySetByMap(Map<String, String> map) {

        Set<String> keySet = new HashSet<String>();
        keySet.addAll(map.keySet());

        return keySet;
    }

    /**
     * 根据key的set返回key的list
     *
     * @param set
     * @return
     */
    public static List<String> getKeyListBySet(Set<String> set) {
        List<String> keyList = new ArrayList<String>();
        keyList.addAll(set);
        return keyList;
    }

    /**
     * 根据map返回key和value的list
     *
     * @param map
     * @param isKey true为key,false为value
     * @return
     */
    public static List<String> getListByMap(Map<String, String> map,
                                            boolean isKey) {
        List<String> list = new ArrayList<String>();

        Iterator<String> it = map.keySet().iterator();
        while (it.hasNext()) {
            String key = it.next();
            if (isKey) {
                list.add(key);
            } else {
                list.add(map.get(key));
            }
        }

        return list;
    }


    /**
     * 根据map返回key和value的list
     *
     * @param
     * @param map
     * @return
     */
    public static List<String> getKeyListByMap(Map<String, WordPictureVO> map) {
        List<String> list = new ArrayList<String>();

        Iterator<String> it = map.keySet().iterator();
        while (it.hasNext()) {
            String key = it.next();
            list.add(key);
        }

        return list;
    }


    /**
     * 因POI 3.8自带的BUG 导致添加进的图片不显示，只有一个图片框，将图片另存为发现里面的图片是一个PNG格式的透明图片
     * <p>
     * 这里自定义添加图片的方法
     * <p>
     * 往Run中插入图片(解决在word中不显示的问题)
     *
     * @param run
     * @param blipId 图片的id
     * @param id     图片的类型
     * @param width  图片的宽
     * @param height 图片的高
     * @author lgj
     */

    public static void addPictureToRun(XWPFRun run, String blipId, int id, int width, int height) {

        final int EMU = 9525;
        width *= EMU;
        height *= EMU;
        CTInline inline = run.getCTR().addNewDrawing().addNewInline();


        String picXml = "" +

                "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +

                "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +

                "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +

                "         <pic:nvPicPr>" +

                "            <pic:cNvPr id=\"" + id + "\" name=\"Generated\"/>" +

                "            <pic:cNvPicPr/>" +

                "         </pic:nvPicPr>" +

                "         <pic:blipFill>" +

                "            <a:blip r:embed=\"" + blipId + "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" +

                "            <a:stretch>" +

                "               <a:fillRect/>" +

                "            </a:stretch>" +

                "         </pic:blipFill>" +

                "         <pic:spPr>" +

                "            <a:xfrm>" +

                "               <a:off x=\"0\" y=\"0\"/>" +

                "               <a:ext cx=\"" + width + "\" cy=\"" + height + "\"/>" +

                "            </a:xfrm>" +

                "            <a:prstGeom prst=\"rect\">" +

                "               <a:avLst/>" +

                "            </a:prstGeom>" +

                "         </pic:spPr>" +

                "      </pic:pic>" +

                "   </a:graphicData>" +

                "</a:graphic>";


        //CTGraphicalObjectData graphicData = inline.addNewGraphic().addNewGraphicData();

        XmlToken xmlToken = null;

        try {

            xmlToken = XmlToken.Factory.parse(picXml);

        } catch (XmlException xe) {

            xe.printStackTrace();

        }

        inline.set(xmlToken);

        inline.setDistT(0);

        inline.setDistB(0);

        inline.setDistL(0);

        inline.setDistR(0);

        org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D extent = inline.addNewExtent();

        extent.setCx(width);

        extent.setCy(height);

        CTNonVisualDrawingProps docPr = inline.addNewDocPr();

        docPr.setId(id);

        docPr.setName("Picture " + id);

        docPr.setDescr("Generated");

    }


    public void repaceWordAndExcel(XWPFDocument doc, Map<String, String> params, Map<String, WordPictureVO> pictures) throws InvalidFormatException, IOException {

        if (params != null) {
            List<String> keys = new ArrayList<>();
            if (pictures != null) {
                keys = getKeyListByMap(pictures);
            }
            Map<String, String> pidIds = new HashMap<>();
            for (String key : keys) {
                pidIds.put(key, doc.addPictureData(pictures.get(key).getPictureData(), pictures.get(key).getPictureType()));
            }
            replaceInPara(doc, params, pictures, pidIds);
            replaceInTable(doc, params, pictures, pidIds);
        }
    }


}


