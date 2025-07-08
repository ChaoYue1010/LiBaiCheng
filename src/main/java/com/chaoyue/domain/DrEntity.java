package com.chaoyue.domain;

import lombok.Data;

import java.util.List;

/**
 * @Author wcy
 * @Date 2025/5/29 10:00
 * @Description:
 */
@Data
public class DrEntity {
    /**
     * 报告编号（AF
     */
    private String baoGaoBianHao;
    /**
     * 指令编号（AE
     */
    private String zhiLingBianHao;
    /**
     * 工艺卡编号
     */
    private String gongYiKaBianHao;
    /**
     * 桩号（I
     */
    private String zhuangHao;
    /**
     * 坡口型式（H
     */
    private String poKouXingShi;
    /**
     * 焊接方法（E
     */
    private String hanJieFangFa;
    /**
     * 管材规格（F
     */
    private String guanCaiGuiGe;
    /**
     * 探测器型号（AX，要是DY-ZR-220514101填RAPIXX 3NDT WIFI，要是CPP1784022填RAPIXX 2NDT WIFI
     */
    private String tanCeQiXingHao;
    /**
     * 像素尺寸（μm）
     */
    private String xiangSuChiCun;

    /**
     * 探测器规格
     */
    private String tanCeQiGuiGe;

    /**
     * 管电压（F（要是φ1219×18.4填260，要是φ1219×22填270）
     */
    private String guanDianYa;
    /**
     * 总曝光时间
     * （AX列要是CPP1784022的话看F列的壁厚φ1219×18.4或者φ1219×22都填9.2
     * （AX列要是DY-ZR-220514101的话看F列的壁厚φ1219×18.4填8.55，φ1219×22的话填9，φ1219×27.5的话填9.2
     */
    private String zongBaoGuangShiJian;
    /**
     * 一次透照长度
     * （AX列要是CPP1784022填119
     * （AX列要是DY-ZR-220514101填226
     */
    private String yiCiTouZhaoChangDu;
    /**
     * 灰度值范围（AK
     */
    private String huiDuZhiFanWei;
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
