package com.alibaba.excel.util;

import com.alibaba.excel.support.ExcelTypeEnum;
import com.alibaba.excel.write.metadata.holder.WriteWorkbookHolder;
import org.apache.poi.POIXMLProperties;
import org.apache.poi.hpsf.CustomProperties;
import org.apache.poi.hpsf.DocumentSummaryInformation;
import org.apache.poi.hssf.record.crypto.Biff8EncryptionKey;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.Map;

/**
 *
 * @author jipengfei
 */
public class WorkBookUtil {

    private static final int ROW_ACCESS_WINDOW_SIZE = 500;

    private WorkBookUtil() {}

    public static void createWorkBook(WriteWorkbookHolder writeWorkbookHolder) throws IOException {
        if (ExcelTypeEnum.XLSX.equals(writeWorkbookHolder.getExcelType())) {
            if (writeWorkbookHolder.getTempTemplateInputStream() != null) {
                XSSFWorkbook xssfWorkbook = new XSSFWorkbook(writeWorkbookHolder.getTempTemplateInputStream());
                putWorkProperties(xssfWorkbook, writeWorkbookHolder.getWriteWorkbook().getWorkProperties());
                writeWorkbookHolder.setCachedWorkbook(xssfWorkbook);
                if (writeWorkbookHolder.getInMemory()) {
                    writeWorkbookHolder.setWorkbook(xssfWorkbook);
                } else {
                    writeWorkbookHolder.setWorkbook(new SXSSFWorkbook(xssfWorkbook, ROW_ACCESS_WINDOW_SIZE));
                }
                return;
            }
            Workbook workbook = null;
            if (writeWorkbookHolder.getInMemory()) {
                workbook = new XSSFWorkbook();
                putWorkProperties(workbook, writeWorkbookHolder.getWriteWorkbook().getWorkProperties());
            } else {
                workbook = new SXSSFWorkbook(ROW_ACCESS_WINDOW_SIZE);
                putWorkProperties(workbook, writeWorkbookHolder.getWriteWorkbook().getWorkProperties());
            }
            writeWorkbookHolder.setCachedWorkbook(workbook);
            writeWorkbookHolder.setWorkbook(workbook);
            return;
        }
        HSSFWorkbook hssfWorkbook;
        if (writeWorkbookHolder.getTempTemplateInputStream() != null) {
            hssfWorkbook = new HSSFWorkbook(new POIFSFileSystem(writeWorkbookHolder.getTempTemplateInputStream()));
        } else {
            hssfWorkbook = new HSSFWorkbook();
        }
        writeWorkbookHolder.setCachedWorkbook(hssfWorkbook);
        writeWorkbookHolder.setWorkbook(hssfWorkbook);
        if (writeWorkbookHolder.getPassword() != null) {
            Biff8EncryptionKey.setCurrentUserPassword(writeWorkbookHolder.getPassword());
            hssfWorkbook.writeProtectWorkbook(writeWorkbookHolder.getPassword(), StringUtils.EMPTY);
        }
    }

    /**
     * 设置文档属性
     */
    private static void putWorkProperties(Workbook workbook, Map<String, Object> properties){
        if (properties == null){
            return;
        }
        if (workbook instanceof HSSFWorkbook){
            HSSFWorkbook hssfWorkbook = (HSSFWorkbook)workbook;
            DocumentSummaryInformation info = hssfWorkbook.getDocumentSummaryInformation();
            CustomProperties customProperties = new CustomProperties();
            for (String key:properties.keySet()){
                customProperties.put(key, properties.get(key));
            }
            info.setCustomProperties(customProperties);
        } else if (workbook instanceof XSSFWorkbook){
            XSSFWorkbook xssfWorkbook = (XSSFWorkbook)workbook;
            POIXMLProperties.CustomProperties customProperties = xssfWorkbook.getProperties().getCustomProperties();
            for (String key:properties.keySet()){
                customProperties.addProperty(key, String.valueOf(properties.get(key)));
            }
        } else if (workbook instanceof SXSSFWorkbook){
            SXSSFWorkbook sxssfWorkbook = (SXSSFWorkbook)workbook;
            POIXMLProperties.CustomProperties customProperties = sxssfWorkbook.getXSSFWorkbook().getProperties().getCustomProperties();
            for (String key:properties.keySet()){
                customProperties.addProperty(key, String.valueOf(properties.get(key)));
            }
        } else {

        }
    }

    public static Sheet createSheet(Workbook workbook, String sheetName) {
        return workbook.createSheet(sheetName);
    }

    public static Row createRow(Sheet sheet, int rowNum) {
        return sheet.createRow(rowNum);
    }

    public static Cell createCell(Row row, int colNum) {
        return row.createCell(colNum);
    }

    public static Cell createCell(Row row, int colNum, CellStyle cellStyle) {
        Cell cell = row.createCell(colNum);
        cell.setCellStyle(cellStyle);
        return cell;
    }

    public static Cell createCell(Row row, int colNum, CellStyle cellStyle, String cellValue) {
        Cell cell = createCell(row, colNum, cellStyle);
        cell.setCellValue(cellValue);
        return cell;
    }

    public static Cell createCell(Row row, int colNum, String cellValue) {
        Cell cell = row.createCell(colNum);
        cell.setCellValue(cellValue);
        return cell;
    }
}
