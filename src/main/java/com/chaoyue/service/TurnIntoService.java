package com.chaoyue.service;

import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

/**
 * @Author wcy
 * @Date 2025/5/19 15:34
 * @Description:
 */
public interface TurnIntoService {

    String toAUT(MultipartFile file, Integer sheetAt, Integer startLine, Integer endLine, HttpServletRequest request, HttpServletResponse response);

    String toDR(MultipartFile file, Integer sheetAt, Integer startLine, Integer endLine, HttpServletRequest request, HttpServletResponse response);

    String toPA(MultipartFile file, Integer sheetAt, Integer startLine, Integer endLine, HttpServletRequest request, HttpServletResponse response);

}
