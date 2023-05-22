package com.mason;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.builder.ExcelWriterBuilder;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import com.alibaba.excel.write.style.column.LongestMatchColumnWidthStyleStrategy;
import com.mason.listener.ExcelListener;
import org.apache.commons.fileupload.disk.DiskFileItem;
import org.apache.commons.fileupload.disk.DiskFileItemFactory;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.HorizontalAlignment;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

/**
 * easyexcel工具类
 *
 * @author gtao
 * @date 2022/8/17 09:44
 * @since 1.0.0
 */
public class ExcelUtils {

    /**
     * 导出
     *
     * @param objects 导出的数据
     * @param clazz 导出的对象类型
     * @param fileName 文件名
     * @throws IOException
     */
    public static void exportExcel(List objects, Class clazz, String fileName) throws IOException {
        exportExcel(objects, clazz, null, fileName);
    }


    /**
     * 错误数据导出请使用本方法
     * 
     * @param objects 导出的数据
     * @param clazz 导出的对象类型
     * @param errMsgList 错误信息（包含错误单元格索引及批注信息）
     * @param fileName 文件名
     * @throws IOException
     */
    public static void exportExcel(List objects, Class clazz, List<Map<Integer, String>> errMsgList, String fileName) throws IOException {
        exportExcel(objects, clazz, errMsgList, null, fileName, "sheet1");
    }


    /**
     * 导出
     * 
     * @param objects 导出的数据
     * @param clazz 导出的对象类型
     * @param errMsgList 错误信息（包含错误单元格索引及批注信息）
     * @param excelMergeDto 合并单元格信息
     * @param fileName 文件名
     * @param sheetName 
     * @throws IOException
     */
    public static void exportExcel(List objects, Class clazz, List<Map<Integer, String>> errMsgList, ExcelMergeDto excelMergeDto, String fileName, String sheetName) throws IOException {
        fileName = fileName + ExcelTypeEnum.XLSX.getValue();
        DiskFileItem fileItem = (DiskFileItem) new DiskFileItemFactory().createItem("file",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", true, fileName);
        // 1设置表头样式
        WriteCellStyle headStyle = new WriteCellStyle();
        // 1.1设置表头数据居中
        headStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        // 2设置表格内容样式
        WriteCellStyle bodyStyle = new WriteCellStyle();
        // 2.1设置表格内容水平居中
        bodyStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        try {
            ExcelWriterBuilder write = EasyExcel.write(fileItem.getOutputStream(), clazz).excelType(
                    ExcelTypeEnum.XLSX).autoCloseStream(Boolean.TRUE)
                    .registerWriteHandler(new LongestMatchColumnWidthStyleStrategy())   //表头自适应长宽
                    .registerWriteHandler(new HorizontalCellStyleStrategy(headStyle, bodyStyle)); //数据居中

            if (excelMergeDto != null) {
                write.registerWriteHandler(new ExcelMergeCellHandler(excelMergeDto.getMergeColIndex(), excelMergeDto.getMergeRowIndex()));
            }

            if (errMsgList != null) {
                //inMemory(Boolean.TRUE)开启批注   批注在ErrorSheetWriteHandler中实现
                write.inMemory(Boolean.TRUE)
                        .registerWriteHandler(new ErrorSheetWriteHandler(errMsgList));
            }
            write.sheet(sheetName).doWrite(objects);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    /**
     * 导入，默认导出错误数据
     * 
     * @param fileInputStream 文件输入流
     * @param excelManage 业务逻辑处理接口
     * @param clazz 对应的DTO对象类型
     * @return ExcelListener
     * @throws Exception
     */
    public static ExcelListener importExcel(InputStream fileInputStream, ExcelManage excelManage, Class clazz) throws Exception {
        return importExcel(fileInputStream, excelManage, clazz, true);
    }

    /**
     * 导入
     *
     * @param fileInputStream 文件输入流
     * @param excelManage     业务逻辑处理接口
     * @param clazz           对应的DTO对象类型
     * @param isErrorExport   是否导出错误数据 默认true
     * @return ExcelListener
     */
    public static ExcelListener importExcel(InputStream fileInputStream, ExcelManage excelManage, Class clazz, boolean isErrorExport) throws Exception {
        ExcelListener excelListener = new ExcelListener(excelManage, clazz, isErrorExport);
        ZipSecureFile.setMinInflateRatio(-1.0d);
        EasyExcel.read(fileInputStream, clazz, excelListener).autoCloseStream(Boolean.TRUE).ignoreEmptyRow(false).sheet().doRead();
        return excelListener;
    }

}
