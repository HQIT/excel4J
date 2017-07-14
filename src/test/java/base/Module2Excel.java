package base;


import com.github.ExcelUtils;
import com.github.sink.ExcelFileSink;
import com.github.source.ExcelFileSource;

import moudles.Student1;

import org.junit.Test;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class Module2Excel {

    @Test
    public void testObject2Excel() throws Exception {
    	
        String tempPath = "C:\\Users\\Administrator\\Documents\\excel4J\\src\\test\\resource\\normal_template.xlsx";
        
        List<Student1> list = new ArrayList<>();
        list.add(new Student1("1010001", "盖伦", "六年级三班"));
        list.add(new Student1("1010002", "古尔丹", "一年级三班"));
        list.add(new Student1("1010003", "蒙多(被开除了)", "六年级一班"));
        list.add(new Student1("1010004", "萝卜特", "三年级二班"));
        list.add(new Student1("1010005", "奥拉基", "三年级二班"));
        list.add(new Student1("1010006", "得嘞", "四年级二班"));
        list.add(new Student1("1010007", "瓜娃子", "五年级一班"));
        list.add(new Student1("1010008", "战三", "二年级一班"));
        list.add(new Student1("1010009", "李四", "一年级一班"));
        Map<String, String> data = new HashMap<>();
        data.put("title", "战争学院花名册");
        data.put("info", "学校统一花名册");
        
        List<Student1> list1 = new ArrayList<>();
        list1.add(new Student1("1010001", "盖伦", "六年级三班"));
        list1.add(new Student1("1010002", "古gubao", "一年级三班"));
        list1.add(new Student1("1010003", "蒙多(被开)", "六年级一班"));
        list1.add(new Student1("1010004", "萝卜特", "三年级二班"));
        list1.add(new Student1("1010005", "基", "三年级二班"));
        list1.add(new Student1("1010006", "得嘞", "四年级二班"));
        list1.add(new Student1("1010007", "瓜娃子", "五年级一班"));
        list1.add(new Student1("1010008", "战三", "二年级一班"));
        list1.add(new Student1("1010009", "李四", "一年级一班"));
        String[] sheetNames = {"一班","二班"};
        
        
        List<List<Student1>> listall = new ArrayList<>();
        listall.add(list);
        listall.add(list1);
        
        // 基于模板导出Excel
        ExcelUtils.getInstance().exportObjects2Excel(ExcelFileSource.create(tempPath), 0, sheetNames, list, data, Student1.class, false, ExcelFileSink.create("A.xlsx"));
        
        // 不基于模板导出Excel
        ExcelUtils.getInstance().exportObjects2Excel(listall, Student1.class, true, sheetNames, true, ExcelFileSink.create("B.xlsx"));
    }

    @Test
    public void testMap2Excel() throws Exception {

        Map<String, List<?>> classes = new HashMap<>();

        Map<String, String> data = new HashMap<>();
        data.put("title", "战争学院花名册");
        data.put("info", "学校统一花名册");

        classes.put("class_one", Arrays.asList(
        		new Student1("1010009", "李四", "一年级一班"),
        		new Student1("1010002", "古尔丹", "一年级三班")
        		));
        classes.put("class_two", Arrays.asList(
        		new Student1("1010008", "战三", "二年级一班")
        		));
        classes.put("class_three", Arrays.asList(
	            new Student1("1010004", "萝卜特", "三年级二班"),
	            new Student1("1010005", "奥拉基", "三年级二班")
	            ));
        classes.put("class_four", Arrays.asList(
        		new Student1("1010006", "得嘞", "四年级二班")
        		));
        classes.put("class_six", Arrays.asList(
	            new Student1("1010001", "盖伦", "六年级三班"),
	            new Student1("1010003", "蒙多", "六年级一班")
        		));
        String[] sheetNames = {"一班","二班"};
        ExcelUtils.getInstance().exportObject2Excel(ExcelFileSource.create("C:\\Users\\Administrator\\Documents\\excel4J\\src\\test\\resource\\map_template.xlsx"),
                0, sheetNames, classes, data, Student1.class, false, ExcelFileSink.create("C.xlsx"));
    }

    @Test
    public void testList2Excel() throws Exception {
    	/*
        List<List<String>> list2 = new ArrayList<>();
        List<String> header = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            List<String> _list = new ArrayList<>();
            for (int j = 0; j < 10; j++) {
                _list.add(i + " -- " + j);
            }
            list2.add(_list);
            header.add(i + "---");
        }
        String[] sheetNames = {"一班","二班"};
        ExcelUtils.getInstance().exportObjects2Excel(list2, header,sheetNames,true, ExcelFileSink.create("D.xlsx"));
        */
    	
    	List<List<String>> list1 = new ArrayList<>();
        List<String> header1 = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            List<String> _list = new ArrayList<>();
            for (int j = 0; j < 10; j++) {
                _list.add(i + " -- " + j);
            }
            list1.add(_list);
            header1.add(i + "---");
        }
    	
    	
    	List<List<String>> list2 = new ArrayList<>();
        List<String> header2 = new ArrayList<>();
        for (int i = 0; i < 10; i++) {
            List<String> _list = new ArrayList<>();
            for (int j = 0; j < 10; j++) {
                _list.add(i + " -- " + j);
            }
            list2.add(_list);
            header2.add(i + "-");
        }
        
        List<List<List<String>>> listAll = new ArrayList<>();
        listAll.add(list1);
        listAll.add(list2);
        
        List<List<String>> headerAll = new ArrayList<>();
        headerAll.add(header1);
        headerAll.add(header2);
        
        String[] sheetNames = {"一班","二班"};
        ExcelUtils.getInstance().exportObjects2Excel(listAll, headerAll,sheetNames,true, ExcelFileSink.create("D.xlsx"));
    	
    	
    }
    
}
