package com.lz.paperword;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.properties.ConfigurationPropertiesScan;

@SpringBootApplication
@ConfigurationPropertiesScan
public class PaperToWordApplication {

    public static void main(String[] args) {
        SpringApplication.run(PaperToWordApplication.class, args);
    }
}
