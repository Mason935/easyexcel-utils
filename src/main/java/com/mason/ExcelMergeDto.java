package com.mason;

/**
 * 单元格合并信息
 *
 * @author gtao
 * @date 2022/8/17 15:08
 * @since 1.0.0
 */
public class ExcelMergeDto {
    // 需要合并的列，从0开始算
    private int[] mergeColIndex;
    // 从指定的行开始合并，从0开始算
    private int mergeRowIndex;

    public ExcelMergeDto(int[] mergeColIndex, int mergeRowIndex) {
        this.mergeColIndex = mergeColIndex;
        this.mergeRowIndex = mergeRowIndex;
    }

    public void setMergeColIndex(int[] mergeColIndex) {
        this.mergeColIndex = mergeColIndex;
    }

    public void setMergeRowIndex(int mergeRowIndex) {
        this.mergeRowIndex = mergeRowIndex;
    }

    public int[] getMergeColIndex() {
        return mergeColIndex;
    }

    public int getMergeRowIndex() {
        return mergeRowIndex;
    }
}
