package com.lz.easyexcel.domain;

import com.alibaba.excel.annotation.ExcelProperty;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Date;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class ExportVo {

    @ExcelProperty(value = "userId", order = 1)
    private String userId;

    @ExcelProperty(value = "userName", order = 2)
    private String userName;
}
