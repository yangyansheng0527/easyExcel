package com.yys.easyexcel.service.impl;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.read.builder.ExcelReaderBuilder;
import com.alibaba.excel.read.builder.ExcelReaderSheetBuilder;
import com.yys.easyexcel.domain.Student;
import com.yys.easyexcel.listener.StudentReadListener;
import com.yys.easyexcel.service.ExcelReadService;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

@Service
@Slf4j
public class ExcelReadServiceImpl implements ExcelReadService {

    @Autowired
    private StudentReadListener studentReadListener;
    @Override
    public void readExcelData(List<Student> studentList) {
        for (Student student : studentList) {
            System.out.println("student"+student);
        }
    }

    @Override
    public void read(MultipartFile exceFile) {
        try {
            InputStream inputStream = exceFile.getInputStream();
            ExcelReaderBuilder excelRead = EasyExcel.read(inputStream, Student.class, studentReadListener);
            ExcelReaderSheetBuilder sheet = excelRead.sheet();
            sheet.doRead();
        } catch (IOException e) {
            log.info(e.getMessage());
        }
    }

}
