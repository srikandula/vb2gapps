package com.infy.gcoe;

import java.util.Arrays;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.EnableAutoConfiguration;
import org.springframework.context.ConfigurableApplicationContext;
import org.springframework.context.annotation.ComponentScan;
import org.springframework.context.annotation.Configuration;

@Configuration
@EnableAutoConfiguration()
@ComponentScan
public class Application {

    private static Logger logger = LoggerFactory.getLogger(Application.class);
    

    public static void main(String[] args) {
        logger.debug("Input args: " + Arrays.toString(args));
        runAppAndLogResult(args);
        System.clearProperty("com.infy.poi.flushSize");
    }

    private static void runAppAndLogResult(String[] args) {
        final ConfigurableApplicationContext run = SpringApplication.run(Application.class, args);
    }

}
