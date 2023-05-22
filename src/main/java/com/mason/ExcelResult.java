package com.mason;

import java.util.ArrayList;
import java.util.List;

/**
 * 返回结果集
 *
 * @author gtao
 * @date 2022/8/17 12:50
 * @since 1.0.0
 */
public class ExcelResult<T> {

    private Integer fail;
    /**
     * 错误数据结果集
     */
    private List<ExcelImportErrDto> errDtos;
    
    public ExcelResult() {
        this.errDtos = new ArrayList<>();
    }
    
    public ExcelResult(List<ExcelImportErrDto> errDtos){
        this.errDtos = errDtos;
    }

    public ExcelResult(List<ExcelImportErrDto> errDtos , Integer fail){
        this.errDtos = errDtos;
        this.fail = fail;
    }

    public List<ExcelImportErrDto> getErrDtos() {
        return errDtos;
    }

    public void setErrDtos(List<ExcelImportErrDto> errDtos) {
        this.errDtos = errDtos;
    }

    public Integer getFail() {
        return fail;
    }
}
