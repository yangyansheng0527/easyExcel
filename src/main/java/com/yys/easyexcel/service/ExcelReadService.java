package com.yys.easyexcel.service;

import com.yys.easyexcel.domain.Student;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

public interface ExcelReadService {
    /**
     * 读取学生信息
     * @param studentList
     */
    public void readExcelData(List<Student> studentList);

    void read(MultipartFile exceFile);

}
