package com.yys.easyexcel.domain;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.annotation.write.style.ColumnWidth;
import com.alibaba.excel.annotation.write.style.ContentRowHeight;
import com.alibaba.excel.annotation.write.style.HeadRowHeight;
import lombok.Data;

import java.util.Date;

@Data
@ContentRowHeight(20) //内容的高度
@HeadRowHeight(50) //表头高
public class Student {
    @ExcelProperty(value = "学生姓名",index = 1)
    @ColumnWidth(20)
    private String studenName;  //学生姓名
    @ExcelProperty(value = "性别",index = 0)
    @ColumnWidth(20)
    private String general; //性别
    @ExcelProperty(value = "出生日期",index = 2)
    @ColumnWidth(30)
    private Date birthday; //出生日期
}
