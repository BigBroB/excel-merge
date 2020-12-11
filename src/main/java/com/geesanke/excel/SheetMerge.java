package com.geesanke.excel;

import com.google.common.collect.Maps;
import lombok.Data;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

/**
 * SheetMerge1
 *
 * @author yeehaw
 * @Date 2020/12/11 10:48
 * @Description
 */
@Data
public class SheetMerge<E> {


    /**
     * 合并格子
     */
    private Map<String, List<CellRangeAddress>> ranges = Maps.newHashMap();
    /**
     * 数据
     */
    private List<E> datum;


    public SheetMerge(List<E> datum) {
        this.datum = datum;
    }

    /**
     * 计算格子
     *
     * @param sheetName
     * @param key
     * @param comparator
     * @param cols
     * @param startRow
     */
    public void range(String sheetName, Function<E, String> key, Comparator<? super String> comparator, int[] cols, int startRow) {
        MergeGroup group = MergeGroup.init(datum, key, comparator, cols, startRow);
        List<CellRangeAddress> cellRangeAddresses = group.computeRange(startRow, cols);

        List<CellRangeAddress> addresses = ranges.get(sheetName);
        if (addresses == null || addresses.isEmpty()) {
            ranges.put(sheetName, cellRangeAddresses);
        } else {
            addresses.addAll(cellRangeAddresses);
            ranges.put(sheetName, addresses);
        }
    }

    /**
     * 合并
     *
     * @param workbook
     */
    public void merge(Workbook workbook) {
        for (String sheetName : ranges.keySet()) {
            Sheet sheet = workbook.getSheet(sheetName);
            List<CellRangeAddress> addresses = ranges.get(sheetName);
            for (CellRangeAddress cellAddresses : addresses) {
                sheet.addMergedRegion(cellAddresses);
            }
        }
    }

}
