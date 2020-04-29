package com.kyle.excel;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import com.kyle.excel.metadata.ExcelBeanValidator;
import com.kyle.excel.metadata.ExcelCellProperty;
import com.kyle.excel.util.FileUtil;
import org.apache.commons.beanutils.ConvertUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.fasterxml.jackson.annotation.JsonInclude;
import com.fasterxml.jackson.core.JsonParseException;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.JsonMappingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;

public class ExcelSugar {
    private static final String EXCEL_XLS = "xls";
    private static final String EXCEL_XLSX = "xlsx";
    private static Workbook workbook = null;

    /**
     * 初始化Workbook工作簿
     *
     * @param fileFullName
     */
    public static void initWorkbook(final String fileFullName) {
        try {
            FileInputStream fileInputStream = new FileInputStream(fileFullName);
            initWorkbook(fileInputStream, FileUtil.getExtension(fileFullName));
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 初始化Workbook工作簿
     *
     * @param inputStream
     * @param fileExtension
     */
    public static void initWorkbook(final InputStream inputStream, String fileExtension) {
        try {
            if (EXCEL_XLS.equals(fileExtension.trim().toLowerCase())) {
                workbook = new HSSFWorkbook(inputStream);
            } else if (EXCEL_XLSX.equals(fileExtension.trim().toLowerCase())) {
                workbook = new XSSFWorkbook(inputStream);
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 将数据写入Excel
     *
     * @param <T>
     * @param fileFullName  文件全路径
     * @param sheetIndex    写数据的sheet页
     * @param startRowIndex 写数据的起始行,从1开始
     * @param data          待写入数据
     * @param headerMap     T的属性Name 和 Excel列头标题 映射集合 (K:T的属性,V:Excel列头标题)
     * @param outputStream 保存的输出流
     *
     * @return outputStream 输出流
     */
    private static <T extends ExcelBeanValidator> OutputStream writeExcel(
            final String fileFullName,
            final int sheetIndex,
            final int startRowIndex,
            final List<T> data,
            final Map<String, String> headerMap,
            OutputStream outputStream) {

        if (!isAllowedFile(fileFullName)) {
            throw new IllegalArgumentException(String.format("File format has to be %s/%s", EXCEL_XLS, EXCEL_XLSX));
        }

        // 转换文件为输入流
        File file = FileUtil.getFile(fileFullName);
        InputStream fileInputStream = FileUtil.openInputStream(file);
        boolean isOffice2003 = FileUtil.isExtensionIgnoreCase(fileFullName, EXCEL_XLS);
        outputStream = outputStream == null ? new ByteArrayOutputStream() : outputStream;
        return writeExcel(fileInputStream, isOffice2003, sheetIndex, startRowIndex, data, headerMap, outputStream);
    }

    /**
     * 将数据写入Excel
     *
     * @param <T>
     * @param inputStream   文件流
     * @param isOffice2003   是否为Office2003 (对于传递的流数据，获取文件类型稍微麻烦，故暂时以传递的方式)
     * @param sheetIndex    写数据的sheet
     * @param startRowIndex 写数据的起始行,从1开始
     * @param datas         待写入数据
     * @param headerMap     T 和 Excel 映射集合 (K:T的属性,V:Excel列头标题)
     * @param outputStream 保存的输出流
     *
     * @return outputStream 输出流
     */
    private static <T extends ExcelBeanValidator> OutputStream writeExcel(
            final InputStream inputStream,
            final boolean isOffice2003,
            final int sheetIndex,
            int startRowIndex,
            final List<T> datas,
            final Map<String, String> headerMap,
            OutputStream outputStream) {

        if(datas == null || datas.size() < 1) {
            throw new IllegalArgumentException("datas cannot be null.");
        }

        try {
            outputStream = outputStream == null ? new ByteArrayOutputStream() : outputStream;

            // 初始化Workbook工作簿
            initWorkbook(inputStream, isOffice2003 ? EXCEL_XLS : EXCEL_XLSX);
            // 读取Sheet
            Sheet sheet = workbook.getSheetAt(sheetIndex);

            // 获取导入数据的Bean中的属性Name 和 Excel列属性 的映射
            Map<String, ExcelColumnBean> beanNameColumnIndexMap = getBeanNameColumnIndexMap(sheet, startRowIndex - 1, datas.get(0), headerMap);

            // 写入数据
            for (T data : datas) {
                if(data == null) continue;

                // 获取数据
                Map<String, Object> dataMap = BeanUtil.objectToMap(data);
                for (Entry<String, Object> dataEntry : dataMap.entrySet()) {
                    if(dataEntry == null) continue;

                    String dataKey = dataEntry.getKey();
                    Object dataValue = dataEntry.getValue();

                    // 获取注解
                    ExcelColumnBean excelColumnBean = beanNameColumnIndexMap.get(dataKey);

                    if(excelColumnBean == null) continue;

                    // 获取列索引
                    int cellnum = excelColumnBean.getIndex();

                    // 获取当前单元格
                    Cell cell = getCell(sheet, startRowIndex, cellnum);

                    // 设置单元格值 和 样式
                    setCellValue(cell, dataValue, excelColumnBean);

                    // 设置列宽度
                    if(startRowIndex == 0 && excelColumnBean.getWith() > 0) {
                        sheet.setColumnWidth(cellnum, excelColumnBean.getWith());
                    }
                }

                startRowIndex++;
            }

            // 保存Excel
            save(outputStream);
        } finally {
            // 关闭workbook
            close();
            // 关闭输入流
            IOUtils.closeQuietly(inputStream);
        }

        return outputStream;
    }

    /**
     * 将数据写入Excel，并保存文件到 localFileFullName
     *
     * @param <T>
     * @param fileFullName  文件全路径
     * @param sheetIndex    写数据的sheet页
     * @param startRowIndex 写数据的起始行,从1开始
     * @param data          待写入数据
     * @param headerMap     T的属性Name 和 Excel列头标题 映射集合 (K:T的属性,V:Excel列头标题)
     * @param localFileFullName 保存本地文件全路径
     */
    public static <T extends ExcelBeanValidator> void saveAsExcel(
            final String fileFullName,
            final int sheetIndex,
            final int startRowIndex,
            final List<T> data,
            final Map<String, String> headerMap,
            final String localFileFullName) {

        ByteArrayOutputStream byteArrayOutputStream = null;
        FileOutputStream fileOutputStream = null;
        try {
            if (!isAllowedFile(fileFullName)) {
                throw new IllegalArgumentException(String.format("File format has to be %s/%s", EXCEL_XLS, EXCEL_XLSX));
            }

            if (!FileUtil.isExistsDic(localFileFullName)) {
                throw new FileNotFoundException("File '" + localFileFullName + "' directory does not exist");
            }

            byteArrayOutputStream = (ByteArrayOutputStream) writeExcel(fileFullName, sheetIndex, startRowIndex, data, headerMap, null);

            fileOutputStream = new FileOutputStream(localFileFullName);
            byteArrayOutputStream.writeTo(fileOutputStream);
            fileOutputStream.flush();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // 关闭输入流
            if(byteArrayOutputStream != null) IOUtils.closeQuietly(byteArrayOutputStream);
            if(fileOutputStream != null) IOUtils.closeQuietly(fileOutputStream);
        }
    }

    /**
     * 将数据写入Excel，并保存文件到 localFileFullName
     *
     * @param <T>
     * @param inputStream   文件流
     * @param isOffice2003   是否为Office2003 (对于传递的流数据，获取文件类型稍微麻烦，故暂时以传递的方式)
     * @param sheetIndex    写数据的sheet页
     * @param startRowIndex 写数据的起始行,从1开始
     * @param data          待写入数据
     * @param headerMap     T的属性Name 和 Excel列头标题 映射集合 (K:T的属性,V:Excel列头标题)
     * @param localFileFullName 保存本地文件全路径
     */
    public static <T extends ExcelBeanValidator> void saveAsExcel(
            final InputStream inputStream,
            final boolean isOffice2003,
            final int sheetIndex,
            final int startRowIndex,
            final List<T> data,
            final Map<String, String> headerMap,
            final String localFileFullName) {

        ByteArrayOutputStream byteArrayOutputStream = null;
        FileOutputStream fileOutputStream = null;
        try {
            if (!FileUtil.isExistsDic(localFileFullName)) {
                throw new FileNotFoundException("File '" + localFileFullName + "' directory does not exist");
            }

            byteArrayOutputStream = (ByteArrayOutputStream) writeExcel(inputStream, isOffice2003, sheetIndex, startRowIndex, data, headerMap, null);

            fileOutputStream = new FileOutputStream(localFileFullName);
            byteArrayOutputStream.writeTo(fileOutputStream);
            fileOutputStream.flush();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // 关闭输入流
            if(byteArrayOutputStream != null) IOUtils.closeQuietly(byteArrayOutputStream);
            if(fileOutputStream != null) IOUtils.closeQuietly(fileOutputStream);
        }
    }

    /**
     * 将数据写入Excel，并保存到输出流
     *
     * @param <T>
     * @param fileFullName  文件全路径
     * @param sheetIndex    写数据的sheet页
     * @param startRowIndex 写数据的起始行,从1开始
     * @param data          待写入数据
     * @param headerMap     T的属性Name 和 Excel列头标题 映射集合 (K:T的属性,V:Excel列头标题)
     * @param outputStream 保存的输出流
     */
    public static <T extends ExcelBeanValidator> void saveAsExcel(
            final String fileFullName,
            final int sheetIndex,
            final int startRowIndex,
            final List<T> data,
            final Map<String, String> headerMap,
            OutputStream outputStream) {

        if (!isAllowedFile(fileFullName)) {
            throw new IllegalArgumentException(String.format("File format has to be %s/%s", EXCEL_XLS, EXCEL_XLSX));
        }

        outputStream = outputStream == null ? new ByteArrayOutputStream() : outputStream;

        writeExcel(fileFullName, sheetIndex, startRowIndex, data, headerMap, outputStream);
    }

    /**
     * 将数据写入Excel，并保存到输出流
     *
     * @param <T>
     * @param inputStream   文件流
     * @param isOffice2003   是否为Office2003 (对于传递的流数据，获取文件类型稍微麻烦，故暂时以传递的方式)
     * @param sheetIndex    写数据的sheet页
     * @param startRowIndex 写数据的起始行,从1开始
     * @param data          待写入数据
     * @param headerMap     T的属性Name 和 Excel列头标题 映射集合 (K:T的属性,V:Excel列头标题)
     * @param outputStream 保存的输出流
     */
    public static <T extends ExcelBeanValidator> void saveAsExcel(
            final InputStream inputStream,
            final boolean isOffice2003,
            final int sheetIndex,
            final int startRowIndex,
            final List<T> data,
            final Map<String, String> headerMap,
            OutputStream outputStream) {

        outputStream = outputStream == null ? new ByteArrayOutputStream() : outputStream;

        writeExcel(inputStream, isOffice2003, sheetIndex, startRowIndex, data, headerMap, outputStream);
    }


    /**
     * 获取指定位置的单元格
     *
     * @param sheet   sheet
     * @param rownum  要获取的行(0开始)
     * @param cellnum 要获取的列(0开始)
     * @return 单元格
     */
    public static Cell getCell(final Sheet sheet, final int rownum, final int cellnum) {
        Row row = sheet.getRow(rownum);
        if(row == null)
            row = sheet.createRow(rownum);
        Cell cell = row.getCell(cellnum);
        if(cell == null)
            cell = row.createCell(cellnum);
        return cell;
    }

    /**
     * 获取指定行的单元格集合
     *
     * @param sheet  sheet
     * @param rownum 要获取的行
     * @return 单元格集合
     */
    public static List<Cell> getCellByRownum(final Sheet sheet, final int rownum) {
        List<Cell> cells = new ArrayList<Cell>();

        Row row = sheet.getRow(rownum);
        // 遍历每单元格记录
        for (Cell cell : row) {
            if (cell != null) {
                cells.add(cell);
            }
        }

        return cells;
    }

    /**
     * 获取指定列的单元格集合
     *
     * @param sheet  sheet
     * @param colnum 要获取的列
     * @return 单元格集合
     */
    public static List<Cell> getCellByColnum(final Sheet sheet, final int colnum) {
        List<Cell> cells = new ArrayList<Cell>(sheet.getLastRowNum());

        // 遍历每行记录
        for (Row row : sheet) {
            // 指定列单元格
            Cell cell = row.getCell(colnum);
            if(cell != null) {
                cells.add(cell);
            }
        }

        return cells;
    }

    /**
     * 获取指定范围的单元格集合
     *
     * @param sheet    sheet
     * @param firstRow 开始行
     * @param lastRow  结束行
     * @param firstCol 开始列
     * @param lastCol  结束列
     * @return 单元格集合
     */
    public static List<Cell> getCellRange(final Sheet sheet, final int firstRow, final int lastRow, final int firstCol,
                                          final int lastCol) {
        if (lastRow < firstRow || lastCol < firstCol) {
            throw new IllegalArgumentException("Invalid cell range, having lastRow < firstRow || lastCol < firstCol, "
                    + "had rows " + lastRow + " >= " + firstRow + " or cells " + lastCol + " >= " + firstCol);
        }

        int height = lastRow - firstRow + 1;
        int width = lastCol - firstCol + 1;
        List<Cell> cells = new ArrayList<Cell>(height * width);

        for (int r = firstRow; r <= lastRow; r++) {
            for (int c = firstCol; c <= lastCol; c++) {
                Row row = sheet.getRow(r);
                if (row == null)
                    sheet.createRow(r);
                Cell cell = row.getCell(c);
                if (cell == null)
                    row.createCell(c);
                cells.add(cell);
            }
        }

        return cells;
    }

    /**
     * 获取单元格 值和位置 的映射
     *
     * @param cell 单元格
     * @return 单元格 值和位置 的映射
     */
    public static Map<String, Integer> getCellValueNumMap(final Cell cell) {
        Map<String, Integer> map = new LinkedHashMap<String, Integer>();

        if (cell == null) {
            return map;
        }

        int columnIndex = cell.getColumnIndex();
        String cellValue = cell.getStringCellValue();
        if (!map.containsKey(cellValue)) {
            map.put(cellValue, columnIndex);
        }

        return map;
    }

    /**
     * 获取单元格集合 值和位置 的映射
     *
     * @param cells 单元格集合
     * @return 单元格集合 值和位置 的映射
     */
    public static Map<String, Integer> getCellValueNumMap(final List<Cell> cells) {
        Map<String, Integer> map = new LinkedHashMap<String, Integer>();

        if (cells == null || cells.size() < 1) {
            return map;
        }

        for (Cell cell : cells) {
            int columnIndex = cell.getColumnIndex();
            String cellValue = cell.getStringCellValue();
            if (!map.containsKey(cellValue)) {
                map.put(cellValue, columnIndex);
            }
        }

        return map;
    }

    /**
     * 获取单元格集合 值、位置、样式 的映射
     *
     * @param cells
     * @return
     */
    public static Map<String, ExcelColumnBean> getCellValueNumStyleMap(final Sheet sheet, final List<Cell> cells) {
        Map<String, ExcelColumnBean> map = new LinkedHashMap<String, ExcelColumnBean>();

        if (cells == null || cells.size() < 1) {
            return map;
        }

        for (Cell cell : cells) {
            int columnIndex = cell.getColumnIndex();
            int rowIndex = cell.getRowIndex() + 1;
            String cellValue = cell.getStringCellValue();
            Cell nextCell = getCell(sheet, rowIndex, columnIndex);
            ExcelColumnBean excelColumnBean = new ExcelColumnBean();
            excelColumnBean.setIndex(columnIndex);
            excelColumnBean.setCellStyle(nextCell.getCellStyle());
            if (!map.containsKey(cellValue)) {
                map.put(cellValue, excelColumnBean);
            }
        }

        return map;
    }


    /**
     * 获取单元格的值
     *
     * @param cell 单元格
     * @return 单元格的值
     */
    public static Object getCellValue(final Cell cell) {
        Object cellValue = null;

        switch (cell.getCellType()) {
            case NUMERIC:
                cellValue = cell.getNumericCellValue();
                break;
            case FORMULA:
                cellValue = String.valueOf(cell.getCellFormula());
                break;
            case BOOLEAN:
                cellValue = cell.getBooleanCellValue();
                break;
            case STRING:
                cellValue = cell.getStringCellValue();
                break;
            default:
                cellValue = StringUtils.EMPTY;
                break;
        }

        return cellValue;
    }

    /**
     * 设置单元格的值
     *
     * @param cell 单元格
     * @param value 单元格的值
     * @param excelColumnBean 值注解
     */
    public static void setCellValue(final Cell cell, Object value, ExcelColumnBean excelColumnBean) {
        // 如果值为null 或者 需要加前后缀，都以String处理
        if(value == null) {
            value = "";
        }
        if(excelColumnBean != null && (!StringUtils.isBlank(excelColumnBean.getPrefix())
                || !StringUtils.isBlank(excelColumnBean.getSuffix()))) {
            String prefix = !StringUtils.isBlank(excelColumnBean.getPrefix()) ? excelColumnBean.getPrefix() : "";
            String suffix = !StringUtils.isBlank(excelColumnBean.getSuffix()) ? excelColumnBean.getSuffix() : "";
            value = prefix + value + suffix;
        }

        // 设置单元格样式
        CellStyle cellStyle = excelColumnBean != null && excelColumnBean.getCellStyle() != null
                ? excelColumnBean.getCellStyle()
                : cell.getCellStyle();

        // 设置单元格值
        if (value instanceof Short) {
            cell.setCellValue(((Short) value));
        } else if (value instanceof Integer) {
            // cell.setCellValue(((Integer) value).intValue());
            cell.setCellValue(((Integer) value));
        } else if (value instanceof Long) {
            cell.setCellValue(((Long) value));
        } else if (value instanceof Float) {
            cell.setCellValue(((Float) value));
        } else if (value instanceof Double) {
            // cell.setCellValue(((Double) value).doubleValue());
            cell.setCellValue(((Double) value));
        } else if (value instanceof Boolean) {
            cell.setCellValue(((Boolean) value));
        } else if (value instanceof Date) {
            // 设置时间格式
            String format = !StringUtils.isBlank(excelColumnBean.getDateFormat())
                    ? excelColumnBean.getDateFormat()
                    : "MM/dd/yyyy HH:mm:ss";
            DataFormat dataFormat = workbook.createDataFormat();
            cellStyle.setDataFormat(dataFormat.getFormat(format));
            cell.setCellValue((Date) value);
        } else {
            cell.setCellValue(ConvertUtils.convert(value));
        }

        // 保存样式
        cell.setCellStyle(cellStyle);
    }

    /**
     * 删除指定列
     *
     * @param sheet sheet
     * @param colnum 要获取的列(0开始)
     */
    public void removeColumn(final Sheet sheet, final int colnum) {
        if (sheet == null) {
            return;
        }

        // 遍历每行记录
        for (Row row : sheet) {
            // 指定列单元格
            Cell cell = row.getCell(colnum);
            if(cell != null) {
                row.removeCell(cell);
            }
        }
    }

    /**
     * 保存Excel
     *
     * @param stream 需要保存的目标流
     */
    public static void save(OutputStream stream) {
        stream = stream != null ? stream : new ByteArrayOutputStream();
        if(workbook != null) {
            try {
                // 保存
                workbook.write(stream);
            } catch (IOException e) {
                throw new RuntimeException(e);
            }
        }
    }

    /**
     * 关闭Excel
     *
     */
    public static void close() {
        if (workbook != null) {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 判断文件是否合法
     *
     * @param fileFullName 文件全路径
     * @return 是否合法
     */
    public static boolean isAllowedFile(final String fileFullName) {
        return FileUtil.isExists(fileFullName)
                && FileUtil.isExtensionIgnoreCase(fileFullName, EXCEL_XLS, EXCEL_XLSX)
                && FileUtil.fileSizeOf(fileFullName) > 0;
    }

    /**
     * 获取导入数据的Bean中的属性Name 和 其注解 的映射
     *
     * @param t 存放数据的Bean
     * @return
     */
    public static <T extends ExcelBeanValidator> Map<String, ExcelColumnBean> getBeanFieldNameAnnotationMap(final T t) {
        Map<String, ExcelColumnBean> map = new LinkedHashMap<String, ExcelColumnBean>();

        Class<? extends ExcelBeanValidator> beanClass = t.getClass();


        // 得到对象所有字段
        Field fields[] = beanClass.getDeclaredFields();

        // 遍历所有字段，对应配置好的字段并赋值
        for (Field field : fields) {
            // 获取字段名称
            String fieldName = field.getName();
            // 获取注解
            ExcelColumnBean excelColumnBean = new ExcelColumnBean();
            if (field.isAnnotationPresent(ExcelColumn.class)) {
                ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                if (excelColumn != null) {
                    excelColumnBean.setName(excelColumn.name());
                    excelColumnBean.setIndex(excelColumn.index());
                    excelColumnBean.setPrefix(excelColumn.prefix());
                    excelColumnBean.setSuffix(excelColumn.suffix());
                    excelColumnBean.setDateFormat(excelColumn.dateFormat());
                    excelColumnBean.setCellStyleJson(excelColumn.cellStyleJson());
                }
            }
            if (!map.containsKey(fieldName)) {
                map.put(fieldName, excelColumnBean);
            }
        }

        return map;
    }

    /**
     * 获取导入数据的Bean中的属性Name、Type、Value 和 其注解 的映射
     *
     * @param t 存放数据的Bean
     * @return
     */
    public static <T extends ExcelBeanValidator> Map<String, ExcelCellBean> getBeanFieldAnnotationMap(final T t) {
        Map<String, ExcelCellBean> map = new LinkedHashMap<String, ExcelCellBean>();

        Class<? extends ExcelBeanValidator> beanClass = t.getClass();

        try {
            ExcelCellBean excelCellBean = new ExcelCellBean();

            // 得到对象所有字段
            Field fields[] = beanClass.getDeclaredFields();

            // 遍历所有字段
            for (Field field : fields) {
                // 获取字段名称
                String fieldName = field.getName();
                excelCellBean.setName(fieldName);
                // 获取字段类型
                Class<?> fieldType = field.getType();
                excelCellBean.setType(fieldType);
                // 抑制Java对修饰符的检查
                field.setAccessible(true);
                // 获取字段值
                Object fieldValue;
                fieldValue = field.get(field);
                excelCellBean.setValue(fieldValue);
                // 获取注解
                ExcelColumnBean excelColumnBean = new ExcelColumnBean();
                if (field.isAnnotationPresent(ExcelColumn.class)) {
                    ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                    if (excelColumn != null) {
                        excelColumnBean.setName(excelColumn.name());
                        excelColumnBean.setIndex(excelColumn.index());
                        excelColumnBean.setPrefix(excelColumn.prefix());
                        excelColumnBean.setSuffix(excelColumn.suffix());
                        excelColumnBean.setDateFormat(excelColumn.dateFormat());
                        excelColumnBean.setCellStyleJson(excelColumn.cellStyleJson());
                    }
                }
                excelCellBean.setExcelColumnBean(excelColumnBean);
                if (!map.containsKey(fieldName)) {
                    map.put(fieldName, excelCellBean);
                }
            }
        } catch (IllegalArgumentException e) {
            e.printStackTrace();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }

        return map;
    }

    /**
     * 获取导入数据的Bean中的属性Name 和 Excel列属性 的映射
     *
     * @param sheet     sheet
     * @param rowIndex  表头行
     * @param t         存放数据的Bean
     * @param headerMap T的属性Name 和 Excel列头标题 映射集合 (K:T的属性,V:Excel列头标题)
     * @return Bean中的属性Name 和 Excel列索引 的映射
     */
    public static <T extends ExcelBeanValidator> Map<String, ExcelColumnBean> getBeanNameColumnIndexMap(
            final Sheet sheet, final int rowIndex, final T t, final Map<String, String> headerMap) {

        Map<String, ExcelColumnBean> map = new LinkedHashMap<String, ExcelColumnBean>();

        // 获取表头
        List<Cell> cells = getCellByRownum(sheet, rowIndex);

        // 获取表头 值、位置、样式 的映射
        Map<String, ExcelColumnBean> sheetHeaderValueNumTypeMap = getCellValueNumStyleMap(sheet, cells);

        // 如果调用者已经指定了：bean属性 和 列头标题 的映射
        if (headerMap != null && headerMap.size() > 0) {
            for (Entry<String, ExcelColumnBean> sheetHeaderEntry : sheetHeaderValueNumTypeMap.entrySet()) {
                if (sheetHeaderEntry == null)
                    continue;

                String sheetHeaderKey = !StringUtils.isBlank(sheetHeaderEntry.getKey())
                        ? sheetHeaderEntry.getKey().trim()
                        : null;
                Integer sheetHeaderValue = sheetHeaderEntry.getValue().getIndex();
                CellStyle sheetHeaderStyle = sheetHeaderEntry.getValue().getCellStyle();

                for (Entry<String, String> headerEntry : headerMap.entrySet()) {
                    if (headerEntry == null)
                        continue;

                    String headerKey = !StringUtils.isBlank(headerEntry.getKey()) ? headerEntry.getKey().trim() : null;
                    String headerValue = !StringUtils.isBlank(headerEntry.getValue()) ? headerEntry.getValue().trim() : null;

                    // 映射
                    if (sheetHeaderKey != null && headerValue != null && sheetHeaderKey.equalsIgnoreCase(headerValue) && !map.containsKey(headerKey)) {
                        ExcelColumnBean excelColumnBean = new ExcelColumnBean();
                        excelColumnBean.setIndex(sheetHeaderValue);
                        excelColumnBean.setCellStyle(sheetHeaderStyle);
                        map.put(headerKey, excelColumnBean);
                    }
                }
            }
        } else { // 通过注解方式设置的属性
            // 获取bean属性 和 bean注解 的映射
            Map<String, ExcelColumnBean> beanAnnotationMap = getBeanFieldNameAnnotationMap(t);

            // 用户设置的索引优先
            if (beanAnnotationMap.values().stream().anyMatch(bean -> bean.getIndex() > -1)) {
                map = beanAnnotationMap;
                // 获取表头样式
                beanAnnotationMap.values().stream().forEach(bean -> {
                    Cell nextCell = getCell(sheet, rowIndex + 1, bean.getIndex());
                    bean.setCellStyle(nextCell.getCellStyle());
                });
            } else {
                for (Entry<String, ExcelColumnBean> sheetHeaderEntry : sheetHeaderValueNumTypeMap.entrySet()) {
                    if (sheetHeaderEntry == null)
                        continue;

                    String sheetHeaderKey = !StringUtils.isBlank(sheetHeaderEntry.getKey())
                            ? sheetHeaderEntry.getKey().trim()
                            : null;
                    Integer sheetHeaderValue = sheetHeaderEntry.getValue().getIndex();
                    CellStyle sheetHeaderStyle = sheetHeaderEntry.getValue().getCellStyle();

                    for (Entry<String, ExcelColumnBean> beanEntry : beanAnnotationMap.entrySet()) {
                        if (beanEntry == null)
                            continue;

                        String beanKey = !StringUtils.isBlank(beanEntry.getKey()) ? beanEntry.getKey().trim() : null;
                        ExcelColumnBean beanValue = beanEntry.getValue();

                        String beanHeaderKey = !StringUtils.isBlank(beanValue.getName()) ? beanValue.getName().trim() : null;

                        // 映射
                        if (sheetHeaderKey != null && beanHeaderKey != null && sheetHeaderKey.equalsIgnoreCase(beanHeaderKey)
                                && !map.containsKey(beanKey)) {
                            ExcelColumnBean excelColumnBean = beanValue != null ? beanValue : new ExcelColumnBean();
                            // 设置索引
                            excelColumnBean.setIndex(sheetHeaderValue);
                            // 如果注解中设置了样式
                            if(!StringUtils.isBlank(beanValue.getCellStyleJson())) {
                                sheetHeaderStyle = convertCellStyleFromJson(beanValue.getCellStyleJson());
                            }
                            excelColumnBean.setCellStyle(sheetHeaderStyle);
                            map.put(beanKey, excelColumnBean);
                            continue;
                        }
                    }
                }
            }
        }

        return map;
    }

    /**
     * 从 CellStyle 字符串转换为 CellStyle
     *
     * @param json
     * @return
     */
    private static CellStyle convertCellStyleFromJson(String json) {
        CellStyle cellStyle = null;

        ObjectMapper objectMapper = new ObjectMapper();
        // 如果为空则不输出
        objectMapper.setSerializationInclusion(JsonInclude.Include.NON_EMPTY);
        // 对于空的对象转json的时候不抛出错误
        objectMapper.disable(SerializationFeature.FAIL_ON_EMPTY_BEANS);
        // 禁用序列化日期为timestamps
        // objectMapper.disable(SerializationFeature.WRITE_DATES_AS_TIMESTAMPS);
        // 禁用遇到未知属性抛出异常
        objectMapper.disable(DeserializationFeature.FAIL_ON_UNKNOWN_PROPERTIES);
        // 视空字符传为null
        objectMapper.enable(DeserializationFeature.ACCEPT_EMPTY_STRING_AS_NULL_OBJECT);

        try {
            cellStyle = objectMapper.readValue(json, CellStyle.class);
        } catch (JsonParseException e) {
            e.printStackTrace();
        } catch (JsonMappingException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return cellStyle;
    }
}
