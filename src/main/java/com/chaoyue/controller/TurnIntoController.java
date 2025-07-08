package com.chaoyue.controller;

import com.chaoyue.service.TurnIntoService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

/**
 * @Author wcy
 * @Date 2025/5/19 15:30
 * @Description:
 */
@RestController
@RequestMapping(value = "/api/v1/chaoYue")
public class TurnIntoController {

    @Autowired
    private TurnIntoService service;

    /**
     * 川气东送二线八标台账转->管道焊缝全自动超声波检测报告
     * @param file 文件
     * @param sheetAt 工作表
     * @param startLine 开始行
     * @param endLine 结束行
     * @param request
     * @param response
     * @return
     */
    @RequestMapping(value = "/toAUT")
    public String toAUT(MultipartFile file, Integer sheetAt, Integer startLine, Integer endLine, HttpServletRequest request, HttpServletResponse response) {
        return service.toAUT(file, sheetAt, startLine, endLine, request, response);
    }

    /**
     * 川气东送二线八标台账转->数字射线检测报告
     * @param file 文件
     * @param sheetAt 工作表
     * @param startLine 开始行
     * @param endLine 结束行
     * @param request
     * @param response
     * @return
     */
    @RequestMapping(value = "/toDR")
    public String toDR(MultipartFile file, Integer sheetAt, Integer startLine, Integer endLine, HttpServletRequest request, HttpServletResponse response) {
        return service.toDR(file, sheetAt, startLine, endLine, request, response);
    }

    /**
     * 川气东送二线八标台账转->相控阵超声检测报告
     * @param file 文件
     * @param sheetAt 工作表
     * @param startLine 开始行
     * @param endLine 结束行
     * @param request
     * @param response
     * @return
     */
    @RequestMapping(value = "/toPA")
    public String toPA(MultipartFile file, Integer sheetAt, Integer startLine, Integer endLine, HttpServletRequest request, HttpServletResponse response) {
        return service.toPA(file, sheetAt, startLine, endLine, request, response);
    }

}
