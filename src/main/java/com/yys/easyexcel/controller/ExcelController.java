package com.yys.easyexcel.controller;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelWriter;
import com.alibaba.excel.enums.WriteDirectionEnum;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.metadata.WriteSheet;
import com.alibaba.excel.write.metadata.fill.FillConfig;
import com.yys.easyexcel.domain.Empl;
import com.yys.easyexcel.domain.Student;
import com.yys.easyexcel.service.ExcelReadService;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.util.ResourceUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;

@RestController
@Slf4j
public class ExcelController {
    @Autowired
    private ExcelReadService excelReadService;
    @RequestMapping(value = "/read",method = RequestMethod.POST)
    public void ExcelRead(MultipartFile exceFile){
        log.info("附件信息"+exceFile);
        excelReadService.read(exceFile);
    }
    @RequestMapping(value = "/write",method = RequestMethod.GET)
    public void ExcelWrite(HttpServletResponse response) throws IOException {
        response.setContentType("application.vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        String fileName = URLEncoder.encode("测试", "utf-8");
        response.setHeader("Content-Disposition","attachment;fileName*=utf-8''"+fileName+".xlsx");
        ExcelWriterBuilder excelWriter = EasyExcel.write(response.getOutputStream(), Student.class);
        List<Student> students = initData();
        excelWriter.sheet().doWrite(students);

    }

    @RequestMapping(value = "/writeBatch",method = RequestMethod.GET)
    public void ExcelWriteBatch(HttpServletResponse response) throws IOException {
        response.setContentType("application.vnd.ms-excel");
        response.setCharacterEncoding("utf-8");
        String fileName = URLEncoder.encode("测试组合", "utf-8");
        response.setHeader("Content-Disposition","attachment;fileName*=utf-8''"+fileName+".xlsx");
        String template="template3.xlsx";
        //按照模板生成工作簿
        ExcelWriter workBook = EasyExcel.write(response.getOutputStream(), Empl.class).withTemplate(template).build();
        //生成工作表
        WriteSheet sheet = EasyExcel.writerSheet().build();
        //在多行数据下的新一行（因多行数据数量不确定）
        FillConfig fillConfig =FillConfig.builder().forceNewRow(true).build();
        //如果是水平填充
        //FillConfig build = FillConfig.builder().direction(WriteDirectionEnum.HORIZONTAL).build();
        //多行数据
        List<Empl> emplList = this.initEmpl();
        //单行数据
        HashMap<String,Object> map =new HashMap<>();
        map.put("number","10");
        map.put("total","100.0");
        workBook.fill(emplList,fillConfig,sheet);
        workBook.fill(map,sheet);
        //关流！！！切记，不然导出的表格打不开
        workBook.finish();

    }

    private List<Student> initData() {
        List<Student> studentList =new ArrayList<Student>();
        for (int i = 0; i < 10; i++) {
            Student student =new Student();
            student.setStudenName("张"+i);
            student.setGeneral("男");
            student.setBirthday(new Date());
             studentList.add(student);
        }
        return studentList;
    }


    private List<Empl> initEmpl() {
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
