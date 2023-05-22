package com.mason;

import java.util.Map;

/**
 * 错误信息封装对象
 *
 * @author gtao
 * @date 2022/8/17 15:51
 * @since 1.0.0
 */
public class ExcelImportErrDto {
    private Object object;

    private Map<Integer,String> cellMap;

    public ExcelImportErrDto(){}

    public ExcelImportErrDto(Object object,Map<Integer,String> cellMap){
        this.object = object;
        this.cellMap = cellMap;
    }

    public Object getObject() {
        return object;
    }

    public void setObject(Object object) {
        this.object = object;
    }

    public Map<Integer, String> getCellMap() {
        return cellMap;
    }

    public void setCellMap(Map<Integer, String> cellMap) {
        this.cellMap = cellMap;
    }
    
}
