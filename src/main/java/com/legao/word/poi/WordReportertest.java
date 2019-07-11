package com.legao.word.poi;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class WordReportertest {
    public static void main(String[] args) throws Exception {

        WordReporter wordReporter = new WordReporter("C:/Users/xiaodao/Desktop/xxx_new.docx");
        Map<String,Object> map = new HashMap<>();
        Map<String ,Object> picMap = new HashMap<>();
        picMap.put("width",333);
        picMap.put("height",222);
        picMap.put("type","jpg");
        picMap.put("path","https://timgsa.baidu.com/timg?image&quality=80&size=b9999_10000&sec=1562837583027&di=58d71a6fc71af77ed98fafdbb379137c&imgtype=0&src=http%3A%2F%2Ff.hiphotos.baidu.com%2Fimage%2Fpic%2Fitem%2Fa71ea8d3fd1f4134d244519d2b1f95cad0c85ee5.jpg");
        map.put("picture",picMap);
        wordReporter.export(map);
        wordReporter.export(map,0);
        List<List<Map<String, Object>>> params = new ArrayList<>();
        for (int i=0;i<100;i++){
            List<Map<String, Object>> tables = new ArrayList<>();
            Map<String, Object> sd = new HashMap<>();

            sd.put("name","gaolei"+i);
            sd.put("age","28"+i);
            sd.put("sex","male"+i);
            sd.put("bir","1992"+i);
            tables.add(sd);
            params.add(tables);
        }
        List<Map<String, Object>> tables = new ArrayList<>();
        Map<String, Object> sd = new HashMap<>();

        sd.put("name","gaolei");
        sd.put("age","28");
        sd.put("sex","male");
        sd.put("bir",picMap);
        tables.add(sd);
        params.add(tables);
        wordReporter.export(params,1,1);
        wordReporter.generate("C:/Users/xiaodao/Desktop/xxx_new2.docx");
    }
}
