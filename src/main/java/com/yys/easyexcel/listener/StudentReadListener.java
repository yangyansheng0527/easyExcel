package com.yys.easyexcel.listener;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.yys.easyexcel.domain.Student;
import com.yys.easyexcel.service.ExcelReadService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.annotation.Scope;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.List;

@Component
@Scope("prototype")
public class StudentReadListener extends AnalysisEventListener<Student> {

    @Autowired
    private ExcelReadService excelReadService;

    List<Student> studentList =new ArrayList<Student>();

    @Override
    public void invoke(Student student, AnalysisContext analysisContext) {
        studentList.add(student);
        if(studentList.size()% studentList.size() ==0){
            excelReadService.readExcelData(studentList);
            studentList.clear();
        }
    }

    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {

    }
}
