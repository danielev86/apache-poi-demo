package com.danielev86.apachepoidemo.service;

import com.danielev86.apachepoidemo.bean.CustomerBean;
import com.danielev86.apachepoidemo.exporter.excel.CustomerExcelExporter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.List;

@Service
public class CustomerService {

    @Autowired
    private CustomerExcelExporter customerExcelExporter;

    public void generateMockupExcel() throws Exception {
        List<CustomerBean> customers = new ArrayList<>();
        customers.add(buildcustomer("Homer", "Simpson", 37));
        customers.add(buildcustomer("Bart", "Simpson", 10));
        XSSFWorkbook wb = customerExcelExporter.createExcelWorkbook(customers);
        customerExcelExporter.exportFile(wb, "/home/daniele/dati.xlsx");
    }

    private CustomerBean buildcustomer(String firstName, String lastName, int age){
        CustomerBean customer = new CustomerBean();
        customer.setFirstName(firstName);
        customer.setLastName(lastName);
        customer.setAge(age);
        return customer;
    }
}
