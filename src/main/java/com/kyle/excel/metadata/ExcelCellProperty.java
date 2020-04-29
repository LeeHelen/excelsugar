package com.kyle.excel.metadata;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 *  Excel列头属性
 *
 * @package: com.kyle.excel.metadata
 * @className: ExcelCellMapper
 * @author: Kyle.Y.Li
 * @since 1.0.0 2020-04-4/29/2020 13:56
 */
public class ExcelCellProperty {
    private String name;
    private int index;
    private int with;
    private String prefix;
    private String suffix;
    private String dateFormat;
    private String cellStyleJson;
    private CellStyle cellStyle;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getIndex() {
        return index;
    }

    public void setIndex(int index) {
        this.index = index;
    }

    public int getWith() {
        return with;
    }

    public void setWith(int with) {
        this.with = with;
    }

    public String getPrefix() {
        return prefix;
    }

    public void setPrefix(String prefix) {
        this.prefix = prefix;
    }

    public String getSuffix() {
        return suffix;
    }

    public void setSuffix(String suffix) {
        this.suffix = suffix;
    }

    public String getDateFormat() {
        return dateFormat;
    }

    public void setDateFormat(String dateFormat) {
        this.dateFormat = dateFormat;
    }

    public String getCellStyleJson() {
        return cellStyleJson;
    }

    public void setCellStyleJson(String cellStyleJson) {
        this.cellStyleJson = cellStyleJson;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }
}
