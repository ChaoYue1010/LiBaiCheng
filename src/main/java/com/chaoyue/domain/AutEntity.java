package com.chaoyue.domain;

import lombok.Data;

import java.util.List;

/**
 * @Author wcy
 * @Date 2025/5/21 20:59
 * @Description:
 */
@Data
public class AutEntity {

    /**
     * 报告编号（N
     */
    private String baoGaoBianHao;
    /**
     * 检测日期（P
     */
    private String jianCeRiQi;
    /**
     * 桩号（I
     */
    private String zhuangHao;
    /**
     * 规格（F
     */
    private String guiGe;
    /**
     * 材质（G
     */
    private String caiZhi;
    /**
     * 坡口型式（H
     */
    private String poKouXingShi;
    /**
     * 焊接方法E
     */
    private String hanJieFangFa;
    /**
     * 设备型号（Y
     */
    private String sheBeiXingHao;
    /**
     * 检测灵敏度（Y
     */
    private Integer jianCeLingMinDu;
    /**
     * 检测数量（N
     */
    private Integer jianCeShuLiang;
    /**
     * 返修数量（V
     */
    private Integer fanXiuShuLiang;
    /**
     * 一次合格率（算的
     */
    private String yiCiHeGeLv;
    /**
     * 合格数量（V
     */
    private Integer heGeShuLiang;
    /**
     * 序号（J 一列多行
     */
    private List<List<String>> xuHao;

}
