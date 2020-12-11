package com.geesanke.excel;

import cn.afterturn.easypoi.excel.annotation.Excel;
import lombok.AllArgsConstructor;
import lombok.Data;

/**
 * TestEntity
 *
 * @author yeehaw
 * @Date 2020/12/11 9:17
 * @Description
 */
@Data
@AllArgsConstructor
public class TestEntity {


    @Excel(name = "A", width = 12, orderNum = "1")
    private String a;

    @Excel(name = "B", width = 12, orderNum = "2")
    private String b;

    @Excel(name = "C", width = 12, orderNum = "3")
    private String c;

    @Excel(name = "D", width = 12, orderNum = "4")
    private String d;

}
