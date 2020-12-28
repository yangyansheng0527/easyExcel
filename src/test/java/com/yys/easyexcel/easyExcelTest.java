package com.yys.easyexcel;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.alibaba.excel.read.builder.ExcelReaderSheetBuilder;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.builder.ExcelWriterSheetBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.sun.org.apache.xml.internal.utils.Hashtree2Node;
import com.yys.easyexcel.domain.Empl;
import com.yys.easyexcel.domain.Student;
import com.yys.easyexcel.listener.StudentListener;
import org.ehcache.xml.model.ResourceUnit;
import org.junit.Test;
import org.springframework.util.ResourceUtils;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.*;

public class easyExcelTest {

    @Test
    public void readData() throws FileNotFoundException {
        //1.获取工作薄
        /**
         * 1.文件路径
         * 2.读取对应的实体
         * 3.监听器
         */
        String path = ResourceUtils.getFile("classpath:templates/学生信息.xlsx").getPath();
        ExcelReaderBuilder excelReader = EasyExcel.read(path, Student.class, new StudentListener());
        //2.获取工作表
        ExcelReaderSheetBuilder sheet = excelReader.sheet();
        //3.获取数据
        sheet.doRead();
    }
    @Test
    public void writeData() throws FileNotFoundException {
        List<Student> studentList =new ArrayList<>();
        Student student1 =new Student();
        student1.setStudenName("张三");
        student1.setGeneral("男");
        student1.setBirthday(new Date("2010/9/10"));
        Student student2 =new Student();
        student2.setStudenName("李四");
        student2.setGeneral("男");
        student2.setBirthday(new Date("2010/9/20"));
        studentList.add(student1);
        studentList.add(student2);
        ExcelWriterBuilder writeBook = EasyExcel.write("学生信息表write.xlsx",Student.class);
        ExcelWriterSheetBuilder sheet = writeBook.sheet();
        sheet.doWrite(studentList);

    }

    @Test
    public void templateWrite(){
        String template="template.xlsx";
        ExcelWriterBuilder excelWriterBuilder = EasyExcel.write("人员信息统计表.xlsx", Empl.class).withTemplate(template);
        ExcelWriterSheetBuilder sheet = excelWriterBuilder.sheet();
        //模拟数据
        Empl empl =new Empl();
        empl.setName("刘小豆");
        empl.setAge("28");
        empl.setHobby("吃饭和睡觉");
        Map<String,Object> map =new HashMap<>(16);
        map.put("name","刘大亮");
        map.put("age","30");
        map.put("hobby","打豆豆");
        sheet.doFill(map);
    }

    @Test
    public void templateBatchWrite(){
        String template="template2.xlsx";
        ExcelWriterBuilder excelWriterBuilder = EasyExcel.write("人员信息统计表(多组).xlsx", Empl.class).withTemplate(template);
        ExcelWriterSheetBuilder sheet = excelWriterBuilder.sheet();
        //模拟数据
        List<Empl> emplList = this.initData();
        sheet.doFill(emplList);
    }

    @Test
    public void templateCompWrite(){
        String template="template3.xlsx";
        //按照模板生成工作簿
        ExcelWriter workBook = EasyExcel.write("人员信息统计表(组合).xlsx", Empl.class).withTemplate(template).build();
        //生成工作表
        WriteSheet sheet = EasyExcel.writerSheet().build();
        //在多行数据下的新一行（因多行数据数量不确定）
        FillConfig fillConfig =FillConfig.builder().forceNewRow(true).build();
        //多行数据
        List<Empl> emplList = this.initData();
        //单行数据
        HashMap<String,Object> map =new HashMap<>();
        map.put("number","10");
        map.put("total","100.0");
        workBook.fill(emplList,fillConfig,sheet);
        workBook.fill(map,sheet);
        //关流！！！
        workBook.finish();
    }
    private List<Empl> initData() {
        List<Empl> emplList =new ArrayList<>();
        for (int i = 0; i < 10; i++) {
             Empl empl =new Empl();
             empl.setName("刘小豆"+i);
             empl.setAge("10"+i);
             empl.setHobby("读书"+i);
             emplList.add(empl);
        }
        return emplList;
    }
}
