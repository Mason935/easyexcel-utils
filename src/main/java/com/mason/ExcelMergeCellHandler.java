package com.mason;

import com.alibaba.excel.metadata.CellData;
import com.alibaba.excel.metadata.Head;
import com.alibaba.excel.write.handler.CellWriteHandler;
import com.alibaba.excel.write.metadata.holder.WriteSheetHolder;
import com.alibaba.excel.write.metadata.holder.WriteTableHolder;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

/**
 * 单元格合并处理Handler类
 *
 * @author gtao
 * @date 2022/8/22 13:43
 * @since 1.0.0
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class ExcelMergeCellHandler implements CellWriteHandler {

    // 需要合并的列，从0开始算
    private int[] mergeColIndex;
    // 从指定的行开始合并，从0开始算
    private int mergeRowIndex;
    
    @Override
    public void beforeCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Row row, Head head, Integer columnIndex, Integer relativeRowIndex, Boolean isHead) {
        
    }

    @Override
    public void afterCellCreate(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {

    }

    @Override
    public void afterCellDataConverted(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, CellData cellData, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {

    }

    /*
     在单元格上的所有操作完成后调用，遍历每一个单元格，判断是否需要向上合并
     */
    @Override
    public void afterCellDispose(WriteSheetHolder writeSheetHolder, WriteTableHolder writeTableHolder, List<CellData> cellDataList, Cell cell, Head head, Integer relativeRowIndex, Boolean isHead) {
        // 获取当前单元格行下标
        int currRowIndex = cell.getRowIndex();
        // 获取当前单元格列下标
        int currColIndex = cell.getColumnIndex();
        // 判断是否大于指定行下标，如果大于则判断列是否也在指定的需要的合并单元列集合中
        if (currRowIndex > mergeRowIndex) {
            for (int i = 0; i < mergeColIndex.length; i++) {
                if (currColIndex == mergeColIndex[i]) {
                    /**
                     * 获取列表数据的唯一标识。不同集合的数据即使数值相同也不合并
                     * 注意：我这里的唯一标识为客户编号（Customer.userCode），在第一列，即下标为0。大家需要结合业务逻辑来做修改
                     */
                    // 获取当前单元格所在的行数据的唯一标识
                    Object currCode = cell.getRow().getCell(0).getNumericCellValue();
                    // 获取当前单元格的正上方的单元格所在的行数据的唯一标识
                    Object preCode = cell.getSheet().getRow(currRowIndex - 1).getCell(0).getNumericCellValue();
                    // 判断两条数据的是否是同一集合，只有同一集合的数据才能合并单元格
                    if(preCode.equals(currCode)){
                        // 如果都符合条件，则向上合并单元格
                        mergeWithPrevRow(writeSheetHolder, cell, currRowIndex, currColIndex);
                        break;
                    }
                }
            }
        }
    }

    /*
     当前单元格向上合并
     */
    private void mergeWithPrevRow(WriteSheetHolder writeSheetHolder, Cell cell, int currRowIndex, int currColIndex) {
        // 获取当前单元格数值
        Object currData = cell.getCellTypeEnum() == CellType.STRING ? cell.getStringCellValue() : cell.getNumericCellValue();
        // 获取当前单元格正上方的单元格对象
        Cell preCell = cell.getSheet().getRow(currRowIndex - 1).getCell(currColIndex);
        // 获取当前单元格正上方的单元格的数值
        Object preData = preCell.getCellTypeEnum() == CellType.STRING ? preCell.getStringCellValue() : preCell.getNumericCellValue();

        // 将当前单元格数值与其正上方单元格的数值比较
        if (preData.equals(currData)) {
            Sheet sheet = writeSheetHolder.getSheet();
            List<CellRangeAddress> mergeRegions = sheet.getMergedRegions();
            // 当前单元格的正上方单元格是否是已合并单元格
            boolean isMerged = false;
            for (int i = 0; i < mergeRegions.size() && !isMerged; i++) {
                CellRangeAddress address = mergeRegions.get(i);
                // 若上一个单元格已经被合并，则先移出原有的合并单元，再重新添加合并单元
                if (address.isInRange(currRowIndex - 1, currColIndex)) {
                    sheet.removeMergedRegion(i);
                    address.setLastRow(currRowIndex);
                    sheet.addMergedRegion(address);
                    isMerged = true;
                }
            }
            // 若上一个单元格未被合并，则新增合并单元
            if (!isMerged) {
                CellRangeAddress cellRangeAddress = new CellRangeAddress(currRowIndex - 1, currRowIndex, currColIndex, currColIndex);
                sheet.addMergedRegion(cellRangeAddress);
            }
        }
    }
    
}
