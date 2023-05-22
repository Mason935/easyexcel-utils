package com.mason.listener;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.enums.CellExtraTypeEnum;
import com.alibaba.excel.event.AnalysisEventListener;
import com.alibaba.excel.exception.ExcelAnalysisException;
import com.alibaba.excel.exception.ExcelDataConvertException;
import com.alibaba.excel.metadata.CellExtra;
import com.alibaba.excel.util.StringUtils;
import com.mason.ExcelImportErrDto;
import com.mason.ExcelManage;
import com.mason.ExcelResult;
import com.mason.ExcelUtils;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.lang.reflect.Field;
import java.text.MessageFormat;
import java.util.*;
import java.util.stream.Collectors;

/**
 * 监听器
 *
 * @author gtao
 * @date 2022/8/17 11:44
 * @since 1.0.0
 */
@Slf4j
public class ExcelListener<T> extends AnalysisEventListener<T> {

    /**
     * 一次性最多导入40000条数据
     */
    private static final int BATCH_COUNT = 40000;

    /**
     * 错误数据结果集
     */
    private List<ExcelImportErrDto> errList = new ArrayList<>();

    /**
     * 处理逻辑service
     */
    private ExcelManage excelManage;

    /**
     * 存放解析的临时数据
     */
    private List<T> list = new ArrayList<>();

    /**
     * 合并单元格信息
     */
    private List<CellExtra> merges = new ArrayList<>();

    /**
     * excel对应的DTO对象的反射类
     */
    private Class<T> clazz;

    /**
     * 是否处理导入错误数据
     */
    private boolean isErrorExport;


    public ExcelListener(ExcelManage excelManage, Class clazz, boolean isErrorExport) {
        this.excelManage = excelManage;
        this.clazz = clazz;
        this.isErrorExport = isErrorExport;
    }

    /**
     * 这个每一条数据解析都会来调用
     *
     * @param data
     * @param context
     */
    @Override
    public void invoke(T data, AnalysisContext context) {
        // 如果一行Excel数据均为空值，则不装载该行数据
        if (isLineNullValue(data)) {
            return;
        }
        list.add(data);
        if (list.size() > BATCH_COUNT) {
            throw new ExcelAnalysisException("数据量超出4万");
        }
    }

    /**
     * 判断整行单元格数据是否均为空
     *
     * @param data
     * @return boolean
     */
    private boolean isLineNullValue(T data) {
        try {
            List<Field> fields = Arrays.stream(data.getClass().getDeclaredFields())
                    .filter(f -> f.isAnnotationPresent(ExcelProperty.class))
                    .collect(Collectors.toList());
            List<Boolean> lineNullList = new ArrayList<>(fields.size());
            for (Field field : fields) {
                field.setAccessible(true);
                Object value = field.get(data);
                //由于我们导入的字段都是String类型
                if (StringUtils.isEmpty((String) value)) {
                    lineNullList.add(Boolean.TRUE);
                } else {
                    lineNullList.add(Boolean.FALSE);
                }
            }
            return lineNullList.stream().allMatch(Boolean.TRUE::equals);
        } catch (Exception e) {
            log.error("读取数据行[{}]解析失败: {}", data, e.getMessage());
        }
        return true;
    }

    /**
     * 读取合并单元格信息
     * 此方法的调用在invoke之后doAfterAllAnalysed之前执行
     *
     * @param extra
     * @param context
     */
    @Override
    public void extra(CellExtra extra, AnalysisContext context) {
        if (CellExtraTypeEnum.MERGE == extra.getType()) {
            merges.add(extra);
        }
    }

    /**
     * 所有数据解析完成了 都会来调用
     *
     * @param context
     */
    @Override
    public void doAfterAllAnalysed(AnalysisContext context) {
        ExcelResult result = excelManage.checkImportExcel(list);
        errList.addAll(result.getErrDtos());
        log.info("数据解析完成！");
        list.clear();
        //这里判断是否导出错误数据
        if (isErrorExport) {
            try {
                exportErrorExcel();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


    /**
     * 在转换异常 获取其他异常下会调用本接口。抛出异常则停止读取。如果这里不抛出异常则 继续读取下一行。
     *
     * @param exception
     * @param context
     * @throws Exception
     */
    @Override
    public void onException(Exception exception, AnalysisContext context) {
        log.error("导入解析失败，但是继续解析下一行:{}", exception.getMessage());
        // 如果是某一个单元格的转换异常 能获取到具体行号
        // 如果要获取头的信息 配合invokeHeadMap使用
        if (exception instanceof ExcelDataConvertException) {
            ExcelDataConvertException excelDataConvertException = (ExcelDataConvertException) exception;
            log.error("导入第{}行，第{}列解析异常，数据为:{}", excelDataConvertException.getRowIndex(),
                    (excelDataConvertException.getColumnIndex() + 1), excelDataConvertException.getCellData());
            throw new ExcelAnalysisException(MessageFormat.format("导入解析数据第{0}行，第{1}列异常，数据为:{2}", excelDataConvertException.getRowIndex(), (excelDataConvertException.getColumnIndex() + 1), excelDataConvertException.getCellData()));
        } else {
            log.error("导入异常", exception.getMessage());
            throw new ExcelAnalysisException("存在不合法的导入数据，详情:" + exception.getMessage());
        }
    }

    /**
     * 这里为一行行的返回头
     *
     * @param headMap
     * @param context
     */
    @SneakyThrows
    @Override
    public void invokeHeadMap(Map<Integer, String> headMap, AnalysisContext context) {
        super.invokeHeadMap(headMap, context);
        if (clazz != null) {
            Map<Integer, String> indexNameMap = getIndexNameMap(clazz);
            Set<Integer> keySet = indexNameMap.keySet();
            for (Integer key : keySet) {
                if (StringUtils.isEmpty(headMap.get(key))) {
                    throw new ExcelAnalysisException("解析excel出错，请上传正确的模板文件");
                }
                if (!headMap.get(key).equals(indexNameMap.get(key))) {
                    throw new ExcelAnalysisException("解析excel出错，请上传正确的模板文件");
                }
            }
        }
    }

    public Map<Integer, String> getIndexNameMap(Class clazz) throws NoSuchFieldException {
        Map<Integer, String> result = new HashMap<>();
        Field field;
        Field[] fields = clazz.getDeclaredFields();
        for (int i = 0; i < fields.length; i++) {
            field = clazz.getDeclaredField(fields[i].getName());
            field.setAccessible(true);
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            if (excelProperty != null) {
                int index = excelProperty.index();
                index = index == -1 ? i : index;
                String[] values = excelProperty.value();
                StringBuilder value = new StringBuilder();
                for (String v : values) {
                    value.append(v);
                }
                result.put(index, value.toString());
            }
        }
        return result;
    }

    /**
     * 错误数据导出
     *
     * @throws IOException
     */
    private void exportErrorExcel() throws IOException {
        //错误数据
        List<T> resultList = errList.stream().map(excelImportErrDto -> {
            return (T) excelImportErrDto.getObject();
        }).collect(Collectors.toList());

        //错误信息（坐标和提示信息）
        List<Map<Integer, String>> errMsgList = errList.stream().map(excelImportErrDto -> {
            return excelImportErrDto.getCellMap();
        }).collect(Collectors.toList());

        Integer failTotle = resultList.size();
        if (failTotle > 0) {
            //调用导出方法
            ExcelUtils.exportExcel(resultList, clazz, errMsgList, "导入错误信息");
        }
    }

}
