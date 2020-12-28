package com.yys.easyexcel.listener;

import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;
import com.yys.easyexcel.domain.Student;
import org.springframework.stereotype.Component;

@Component
public class StudentListener extends AnalysisEventListener<Student> {
    /**
     * 每读取一行数据都会调用一次监听
     * @param student
     * @param analysisContext
     */
    @Override
    public void invoke(Student student, AnalysisContext analysisContext) {
        System.out.println("student:"+student);
    }

    /**
     * 读取完后调用的方法
     * @param analysisContext
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext analysisContext) {

    }
}
