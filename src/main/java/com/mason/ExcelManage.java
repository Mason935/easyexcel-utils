package com.mason;


import java.util.List;

/**
 * 自定义导入的处理逻辑需继承此接口
 *
 * @author gtao
 * @date 2022/8/17 18:44
 * @since 1.0.0
 */
public interface ExcelManage<T> {

    ExcelResult checkImportExcel(List<T> objects);
}
