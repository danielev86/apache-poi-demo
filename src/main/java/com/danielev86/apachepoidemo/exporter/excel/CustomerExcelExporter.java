package com.danielev86.apachepoidemo.exporter.excel;

import com.danielev86.apachepoidemo.bean.CustomerBean;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.util.Arrays;
import java.util.List;

@Component
public class CustomerExcelExporter extends AbstractExcelExporter<List<CustomerBean> >{
    @Override
    public void processData(XSSFWorkbook wb, List<CustomerBean> customers) throws Exception {
        if (CollectionUtils.isNotEmpty(customers)){
            Sheet sheet = createSheet(wb, "First page");
            Font font = wb.createFont();
            font.setItalic(true);
            buildHeaders(wb, sheet);
            int rowNum = 1;
            for (CustomerBean customer : customers){
                int cellNum = 0;
                Row row = createRow(sheet, rowNum);
                createStringCell(wb, row, cellNum++, font, customer.getFirstName());
                createStringCell(wb, row, cellNum++, font, customer.getLastName());
                createNumeric(wb, row, cellNum++, font, customer.getAge());
                rowNum++;
            }
        }
    }

    public void buildHeaders(XSSFWorkbook wb, Sheet sheet) {
        List<String> headers = Arrays.asList("First name", "Last name", "Age");
        Row row = sheet.createRow(0);
        Font font = wb.createFont();
        font.setItalic(true);
        font.setBold(true);
        for (int cellNum = 0; cellNum < headers.size() ; cellNum++){
            createCellHeader(wb, row, cellNum, font, headers.get(cellNum).toUpperCase());
            sheet.setColumnWidth(cellNum, 5000);
        }
    }
}
