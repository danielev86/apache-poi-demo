package com.danielev86.apachepoidemo.exporter.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Date;

@Slf4j
public abstract class AbstractExcelExporter<T extends Object> {

    public void exportFile(XSSFWorkbook wb, String fileName){
        try(OutputStream out = new FileOutputStream(fileName)){
            wb.write(out);
        } catch (Exception e) {
            log.error("Error exporting xsl file!: {}", e.getMessage(), e);
        }
    }

    public Sheet createSheet(XSSFWorkbook wb, String sheetName){
        return wb.createSheet(sheetName);
    }

    public Row createRow(Sheet sheet, int rowNum){
        return sheet.createRow(rowNum);
    }

    public void createCellHeader(XSSFWorkbook wb, Row row, int cellNumber, Font font, String value){
        Cell cell = row.createCell(cellNumber);
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.JUSTIFY);
        cellStyle.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        cell.setCellType(CellType.STRING);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(value);

    }

    public void createStringCell(XSSFWorkbook wb, Row row, int cellNumber, Font font, String value){
        Cell cell = row.createCell(cellNumber);
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.LEFT);
        cell.setCellType(CellType.STRING);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(value);
    }

    public void createImportCell(XSSFWorkbook wb, Row row, int cellNumber, Font font, Number value){
        Cell cell = row.createCell(cellNumber);
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.RIGHT);
        DataFormat dataFormat = createDataFormat(wb);
        cellStyle.setDataFormat(dataFormat.getFormat("0.00"));
        cell.setCellType(CellType.NUMERIC);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(value.doubleValue());
    }

    public void createPercentageCell(XSSFWorkbook wb, Row row, int cellNumber, Font font, Number value){
        Cell cell = row.createCell(cellNumber);
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.RIGHT);
        DataFormat dataFormat = createDataFormat(wb);
        cellStyle.setDataFormat(dataFormat.getFormat("0.00 %"));
        cell.setCellType(CellType.NUMERIC);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(value.doubleValue());
    }

    public void createNumeric(XSSFWorkbook wb, Row row, int cellNumber, Font font, Number value){
        Cell cell = row.createCell(cellNumber);
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.RIGHT);
        cell.setCellType(CellType.NUMERIC);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(value.doubleValue());
    }

    public void createDateCell(XSSFWorkbook wb, Row row, int cellNumber, Font font, Date date){
        Cell cell = row.createCell(cellNumber);
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        DataFormat dataFormat = createDataFormat(wb);
        cellStyle.setDataFormat(dataFormat.getFormat("dd/MM/yyyy"));
        cell.setCellType(CellType.STRING);
        cell.setCellStyle(cellStyle);
        cell.setCellValue(date);
    }

    public XSSFWorkbook createExcelWorkbook(T data)throws Exception{
        XSSFWorkbook wb = new XSSFWorkbook();
        processData(wb, data);
        return wb;
    }

    private DataFormat createDataFormat(XSSFWorkbook wb){
        return wb.createDataFormat();
    }

    public abstract void processData(XSSFWorkbook wb, T data)throws Exception;

}
