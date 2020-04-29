package com.kyle.excel.enums;

import com.kyle.excel.exception.ExcelAnalysisException;
import com.kyle.excel.util.FileUtil;
import java.io.File;

/**
 *
 *
 * @package: com.kyle.excel.enums
 * @className: ExcelTypeEnum
 * @author: Kyle.Y.Li
 * @since 1.0.0 2020-04-4/29/2020 14:18
 */
public enum ExcelTypeEnum {
    /**
     * xls
     */
    XLS(".xls"),
    /**
     * xlsx
     */
    XLSX(".xlsx");

    ExcelTypeEnum(String value) {
        this.setValue(value);
    }

    private String value;

    public String getValue() {
        return value;
    }

    public void setValue(String value) {
        this.value = value;
    }

    public static ExcelTypeEnum valueOf(final File file, int i) {
        if(FileUtil.isExists(file)) {
            throw new ExcelAnalysisException("File does not exist");
        }

        String fileName = file.getName();
        if (fileName.endsWith(XLSX.getValue())) {
            return XLSX;
        } else if (fileName.endsWith(XLS.getValue())) {
            return XLS;
        }

        return null;
    }

    public static ExcelTypeEnum getStatusEnumInstance(String nameOrValue) {
        ExcelTypeEnum[] statusConstants = ExcelTypeEnum.values();

        for (ExcelTypeEnum statusConstant : statusConstants) {
            if(statusConstant.name().equals(nameOrValue) || statusConstant.getValue().equals(nameOrValue)) {
                return statusConstant;
            }
        }

        return null;
    }
}
