package com.chaoyue.config;

import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import javax.servlet.MultipartConfigElement;

/**
 * @Author wcy
 * @Date 2025/5/21 21:57
 * @Description:
 */
@Configuration
public class FileUploadConfig {

    @Bean
    public MultipartConfigElement multipartConfigElement() {
        return new MultipartConfigElement("", 10485760, 52428800, 0);
    }

}