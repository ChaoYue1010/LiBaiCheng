package com.chaoyue.domain;

import lombok.Data;

import java.util.List;

/**
 * @Author wcy
 * @Date 2025/6/27 17:07
 * @Description:
 */
@Data
public class PaEntity {
    /**
     * 报告编号（N
     */
    private String baoGaoBianHao;
    /**
     * 桩号（I
     */
    private String zhuangHao;
    /**
     * 管材规格（F
     */
    private String guanCaiGuiGe;
    /**
     * 坡口型式（H
     */
    private String poKouXingShi;
    /**
     * 焊接方法（E
     */
    private String hanJieFangFa;
    /**
     * 设备编号（AB
     */
    private String sheBeiBianHao;
    /**
     * 焊缝盖帽宽（F
     */
    private String hanFengGaiMaoKuan;
    /**
     * 检测数量（AF
     */
    private Integer jianCeShuLiang;
    /**
     * 返修数量（AV
     */
    private Integer fanXiuShuLiang;
    /**
     * 一次合格率（算的
     */
    private String yiCiHeGeLv;
    /**
     * 合格数量（AV
     */
    private Integer heGeShuLiang;
    /**
     * 检测日期（AH
     */
    private String jianCeRiQi;
    /**
     * 序号（
     */
    private List<List<String>> xuHao;

}
