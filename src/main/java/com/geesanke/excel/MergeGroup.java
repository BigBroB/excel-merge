package com.geesanke.excel;

import com.google.common.collect.Lists;
import lombok.Data;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Comparator;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * MergeGroup
 *
 * @author yeehaw
 * @Date 2020/12/11 9:57
 * @Description
 */
@Data
public class MergeGroup {

    /**
     * 合并组
     */
    private List<Integer> groups;

    /**
     * 和并列
     */
    private int[] cols;

    /**
     * 起始行
     */
    private int startRow;


    /**
     * 计算合并格子
     *
     * @param startRow 表头最后行
     * @param cols     需要合并的列 数组
     * @return
     */
    public List<CellRangeAddress> computeRange(int startRow, int[] cols) {
        List<CellRangeAddress> mergeCell = Lists.newArrayList();
        for (Integer group : this.groups) {
            int interval = group;
            if (interval == 1) {
                startRow = startRow + 1;
                continue;
            }
            int beginRow = startRow;
            int endRow = beginRow + interval - 1;
            for (int col : cols) {
                CellRangeAddress range = singleRowCellRange(beginRow, endRow, col);
                mergeCell.add(range);
            }
            startRow = endRow + 1;
        }
        return mergeCell;
    }

    /**
     * 生成合并单元格 range
     *
     * @param beginRow
     * @param endRow
     * @return
     */
    private CellRangeAddress singleRowCellRange(int beginRow, int endRow, int col) {
        return new CellRangeAddress(beginRow, endRow, col, col);
    }


    /**
     * 初始化一个合并
     *
     * @param list
     * @param key
     * @param comparator
     * @param cols
     * @param startRow
     * @param <E>
     * @return
     */
    public static <E> MergeGroup init(List<E> list, Function<E, String> key, Comparator<? super String> comparator, int[] cols, int startRow) {
        TreeMap<String, List<E>> treeMap = list.stream().collect(Collectors.groupingBy(key, () -> new TreeMap<String, List<E>>(comparator), Collectors.toList()));
        List<Integer> groups = Lists.newArrayList();
        for (Map.Entry<String, List<E>> entry : treeMap.entrySet()) {
            groups.add(entry.getValue().size());
        }
        MergeGroup mergeGroup = new MergeGroup();
        mergeGroup.setCols(cols);
        mergeGroup.setStartRow(startRow);
        mergeGroup.setGroups(groups);
        return mergeGroup;
    }

}
