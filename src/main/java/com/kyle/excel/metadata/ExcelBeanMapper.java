package com.kyle.excel.metadata;

/**
 * Excel数据Bean类的属性 和 excel列头 的映射器
 * <p>建立映射关系，方便写入Excel Cell</p>
 *
 * @package: com.kyle.excel.metadata
 * @className: ExcelBeanMapper
 * @author: Kyle.Y.Li
 * @since 1.0.0 2020-04-4/29/2020 13:58
 */
public class ExcelBeanMapper {
    /**
     * Excel数据Bean类的属性
     */
    private String name;
    private Class<?> type;
    private Object value;
    /**
     * Excel列头属性
     */
    private ExcelCellProperty excelCellProperty;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Class<?> getType() {
        return type;
    }

    public void setType(Class<?> type) {
        this.type = type;
    }

    public Object getValue() {
        return value;
    }

    public void setValue(Object value) {
        this.value = value;
    }

    public ExcelCellProperty getExcelCellProperty() {
        return excelCellProperty;
    }

    public void setExcelCellProperty(ExcelCellProperty excelCellProperty) {
        this.excelCellProperty = excelCellProperty;
    }
}
