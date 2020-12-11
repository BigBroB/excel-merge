package com.geesanke.excel;

import cn.afterturn.easypoi.excel.ExcelExportUtil;
import cn.afterturn.easypoi.excel.entity.ExportParams;
import cn.hutool.core.comparator.CompareUtil;
import cn.hutool.core.io.FileUtil;
import com.google.common.collect.Lists;
import lombok.Cleanup;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.BufferedOutputStream;
import java.io.IOException;
import java.util.List;

/**
 * Test
 *
 * @author yeehaw
 * @Date 2020/12/11 9:23
 * @Description
 */
public class Test {

    public static final String file_name = "C:/Users/xr/Desktop/test/test.xls";

    public static void main(String[] args) throws IOException {
        TestEntity e1 = new TestEntity("a1", "b1", "c1", "d1");
        TestEntity e2 = new TestEntity("a1", "b2", "c2", "d2");
        TestEntity e3 = new TestEntity("a2", "b3", "c3", "d3");
        TestEntity e4 = new TestEntity("a2", "b4", "c4", "d4");
        TestEntity e5 = new TestEntity("a3", "b5", "c5", "d5");
        List<TestEntity> list = Lists.newArrayList(e1, e2, e3, e4, e5);

        SheetMerge<TestEntity> sheetMerge = new SheetMerge(list);
        sheetMerge.range("test", TestEntity::getA, (a, b) -> CompareUtil.compare(a, b), new int[]{0}, 1);

        ExportParams exportParams = new ExportParams();
        exportParams.setSheetName("test");
        Workbook workbook = ExcelExportUtil.exportExcel(exportParams, TestEntity.class, list);
        sheetMerge.merge(workbook);
        @Cleanup
        BufferedOutputStream outputStream = FileUtil.getOutputStream(file_name);
        workbook.write(outputStream);
    }

}
