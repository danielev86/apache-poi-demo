package com.danielev86.apachepoidemo;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication(scanBasePackages = {"com.danielev86"})
public class ApachePoiDemoApplication {

    public static void main(String[] args) {
        SpringApplication.run(ApachePoiDemoApplication.class, args);
    }

}
