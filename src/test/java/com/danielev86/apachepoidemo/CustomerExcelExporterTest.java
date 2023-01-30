package com.danielev86.apachepoidemo;

import com.danielev86.apachepoidemo.bean.CustomerBean;
import com.danielev86.apachepoidemo.exporter.excel.CustomerExcelExporter;
import com.danielev86.apachepoidemo.service.CustomerService;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.ArrayList;
import java.util.List;

@SpringBootTest
public class CustomerExcelExporterTest {

    @Autowired
    private CustomerService customerService;

    @Test
    public void customerExcelTest() throws Exception {
        customerService.generateMockupExcel();
    }

}
