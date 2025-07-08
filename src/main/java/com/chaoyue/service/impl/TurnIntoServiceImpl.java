package com.chaoyue.service.impl;

import com.chaoyue.domain.AutEntity;
import com.chaoyue.domain.DrEntity;
import com.chaoyue.domain.PaEntity;
import com.chaoyue.service.TurnIntoService;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.beans.BeanUtils;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.math.BigInteger;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * @Author wcy
 * @Date 2025/5/19 15:34
 * @Description:
 */
@Slf4j
@Service
public class TurnIntoServiceImpl implements TurnIntoService {

    private static final String ROOT_PATH = System.getProperty("user.dir");
    private static final String AUT_FILE_PATH = ROOT_PATH + File.separator + "word" + File.separator + "AUT.docx";
    private static final String DR_FILE_PATH = ROOT_PATH + File.separator + "word" + File.separator + "DR.docx";
    private static final String PA_FILE_PATH = ROOT_PATH + File.separator + "word" + File.separator + "PA.docx";


    @Override
    public String toAUT(MultipartFile file, Integer sheetAt, Integer startLine, Integer endLine, HttpServletRequest request, HttpServletResponse response) {
        if (file.isEmpty()) {
            return "请选择要上传的文件！";
        }
        if (sheetAt == null) {
            return "请选择要上传的sheet页！";
        }
        if (sheetAt < 1) {
            return "请选择正确的sheet页！";
        }
        if (startLine == null || endLine == null) {
            return "请选择要上传的行数！";
        }
        AutEntity autEntity = ParsingExcelToAut(file, sheetAt, startLine, endLine);
        if (autEntity == null) {
            return "AUT 数据为空！";
        }

        String fileName = autEntity.getBaoGaoBianHao() + ".docx";
        // 根据UA设置文件名编码
        String header = request.getHeader("User-Agent").toUpperCase();
        try {
            if (header.contains("MSIE") || header.contains("TRIDENT") || header.contains("EDGE")) {
                fileName = URLEncoder.encode(fileName, "utf-8").replace("+", "%20"); // IE下载文件名空格变+号问题
            } else {
                fileName = new String(fileName.getBytes(), "ISO8859-1");
            }
        } catch (UnsupportedEncodingException e) {
            log.error("导出文件名编码失败", e);
        }

        response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"");
        response.setContentType("application/octet-stream");
        try (XWPFDocument document = new XWPFDocument(new FileInputStream(AUT_FILE_PATH))) {
            // 2. 获取或创建 Section Properties（控制页面布局）
            CTSectPr sectPr = document.getDocument().getBody().isSetSectPr()
                    ? document.getDocument().getBody().getSectPr()
                    : document.getDocument().getBody().addNewSectPr();
            // 3. 获取或创建页面边距设置
            CTPageMar pageMar = sectPr.isSetPgMar() ? sectPr.getPgMar() : sectPr.addNewPgMar();

            pageMar.setTop(BigInteger.valueOf(1701)); // 设置上边距为3厘米
            pageMar.setLeft(BigInteger.valueOf(1417)); // 设置左边距为2.5厘米
            pageMar.setRight(BigInteger.valueOf(1134)); // 设置右边距为2厘米
            pageMar.setBottom(BigInteger.valueOf(1588));// 设置底边距为2.8厘米
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.createRun();
            changeTableTextAut(document, autEntity);
            OutputStream out = response.getOutputStream();
            document.write(out);
            out.flush();
            out.close();
            System.out.println("导出文件成功");
        } catch (IOException e) {
            log.error("导出文件失败", e);
            return "导出文件失败";
        }
        return "SUCCESS";
    }

    @Override
    public String toDR(MultipartFile file, Integer sheetAt, Integer startLine, Integer endLine, HttpServletRequest request, HttpServletResponse response) {
        if (file.isEmpty()) {
            return "请选择要上传的文件！";
        }
        if (sheetAt == null) {
            return "请选择要上传的sheet页！";
        }
        if (sheetAt < 1) {
            return "请选择正确的sheet页！";
        }
        if (startLine == null || endLine == null) {
            return "请选择要上传的行数！";
        }
        DrEntity drEntity = ParsingExcelToDr(file, sheetAt, startLine, endLine);
        if (drEntity == null) {
            return "DR 数据为空！";
        }

        String fileName = drEntity.getBaoGaoBianHao() + ".docx";
        // 根据UA设置文件名编码
        String header = request.getHeader("User-Agent").toUpperCase();
        try {
            if (header.contains("MSIE") || header.contains("TRIDENT") || header.contains("EDGE")) {
                fileName = URLEncoder.encode(fileName, "utf-8").replace("+", "%20"); // IE下载文件名空格变+号问题
            } else {
                fileName = new String(fileName.getBytes(), "ISO8859-1");
            }
        } catch (UnsupportedEncodingException e) {
            log.error("导出文件名编码失败", e);
        }

        response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"");
        response.setContentType("application/octet-stream");
        try (XWPFDocument document = new XWPFDocument(new FileInputStream(DR_FILE_PATH))) {
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.createRun();
            changeTableTextDr(document, drEntity);
            OutputStream out = response.getOutputStream();
            document.write(out);
            out.flush();
            out.close();
            System.out.println("导出文件成功");
        } catch (IOException e) {
            log.error("导出文件失败", e);
            return "导出文件失败";
        }
        return "SUCCESS";
    }

    @Override
    public String toPA(MultipartFile file, Integer sheetAt, Integer startLine, Integer endLine, HttpServletRequest request, HttpServletResponse response) {
//        if (file.isEmpty()) {
//            return "请选择要上传的文件！";
//        }
//        if (sheetAt == null) {
//            return "请选择要上传的sheet页！";
//        }
//        if (sheetAt < 1) {
//            return "请选择正确的sheet页！";
//        }
//        if (startLine == null || endLine == null) {
//            return "请选择要上传的行数！";
//        }
//        PaEntity paEntity = ParsingExcelToPa(file, sheetAt, startLine, endLine);
//        if (paEntity == null) {
//            return "PA 数据为空！";
//        }
//
//        String fileName = paEntity.getBaoGaoBianHao() + ".docx";
//        // 根据UA设置文件名编码
//        String header = request.getHeader("User-Agent").toUpperCase();
//        try {
//            if (header.contains("MSIE") || header.contains("TRIDENT") || header.contains("EDGE")) {
//                fileName = URLEncoder.encode(fileName, "utf-8").replace("+", "%20"); // IE下载文件名空格变+号问题
//            } else {
//                fileName = new String(fileName.getBytes(), "ISO8859-1");
//            }
//        } catch (UnsupportedEncodingException e) {
//            log.error("导出文件名编码失败", e);
//        }
//
//        response.setHeader("Content-Disposition", "attachment; filename=\"" + fileName + "\"");
//        response.setContentType("application/octet-stream");
//        try (XWPFDocument document = new XWPFDocument(new FileInputStream(DR_FILE_PATH))) {
//            XWPFParagraph paragraph = document.createParagraph();
//            paragraph.createRun();
//            changeTableTextPa(document, paEntity);
//            OutputStream out = response.getOutputStream();
//            document.write(out);
//            out.flush();
//            out.close();
//            System.out.println("导出文件成功");
//        } catch (IOException e) {
//            log.error("导出文件失败", e);
//            return "导出文件失败";
//        }
        return "SUCCESS";
    }


    /**
     * 替换表格内的文字
     *
     * @param document
     * @param autEntity
     */
    private static void changeTableTextAut(XWPFDocument document, AutEntity autEntity) {
        Iterator<XWPFTable> tableIt = document.getTablesIterator();
        while (tableIt.hasNext()) {
            XWPFTable table = tableIt.next();
            if (!checkText(table.getText())) {
                continue;
            }

            List<Integer> replaceIndexes = new ArrayList<>();
            List<XWPFTableRow> rows = table.getRows();

            for (int i = 0; i < rows.size(); i++) {
                XWPFTableRow row = rows.get(i);
                List<XWPFTableCell> cells = row.getTableCells();
                if (cells == null || cells.isEmpty()) continue;

                for (XWPFTableCell cell : cells) {
                    String text = getCellText(cell).trim();
                    if ("${xuHao}".equals(text)) {
                        replaceIndexes.add(i);
                        break;
                    }

                    // 缓存单元格文本，避免重复调用 getText()
                    String cellText = cell.getText();
                    if (checkText(cellText)) {
                        List<XWPFParagraph> paragraphs = cell.getParagraphs();
                        if (paragraphs != null && !paragraphs.isEmpty()) {
                            for (XWPFParagraph paragraph : paragraphs) {
                                replaceValueAut(paragraph, autEntity);
                            }
                        }
                    }
                }
            }

            // 倒序处理，防止插入影响后续索引
            Collections.sort(replaceIndexes, Collections.reverseOrder());

            for (Integer rowIndex : replaceIndexes) {
                List<List<String>> dataList = autEntity.getXuHao();
                if (dataList == null || dataList.isEmpty()) continue;

                int insertIndex = rowIndex + 1; // 在占位符行的下一行插入

                for (List<String> rowData : dataList) {
                    XWPFTableRow newRow = table.insertNewTableRow(insertIndex++);
                    // 设置行高0.85厘米
                    newRow.setHeight(482);

                    for (String cellValue : rowData) {
                        XWPFTableCell cell = newRow.addNewTableCell();
                        cell.removeParagraph(0); // 删除默认段落
                        XWPFParagraph para = cell.addParagraph();
                        para.setAlignment(ParagraphAlignment.CENTER);
                        para.setSpacingBefore(0);// 段前间距
                        para.setSpacingAfter(0);// 段后间距
                        XWPFRun run = para.createRun();
                        run.setText(cellValue);
                        run.setFontSize(10.5); // 设置为5号字体
                        run.setFontFamily("宋体");
                        // 设置垂直居中
                        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                        //设置行间距为12磅（1磅=20 twip，12磅=240 twip）
                        CTP ctp = para.getCTP();
                        CTPPr ppr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
                        CTSpacing spacing = ppr.isSetSpacing() ? ppr.getSpacing() : ppr.addNewSpacing();
                        spacing.setLineRule(STLineSpacingRule.EXACT);
                        spacing.setLine(BigInteger.valueOf(240));
                    }

                    // 合并单元格逻辑...
                    if (rowData.size() > 1) {
                        mergeCellsHorizontal(newRow, 1, 3);
                        mergeCellsHorizontal(newRow, 2, 3);
                        mergeCellsHorizontal(newRow, 3, 4);
                        mergeCellsHorizontal(newRow, 4, 5);
                    }
                }

                // 可选：插入完成后删除原始占位符行
                table.removeRow(rowIndex);

                // 然后从插入内容的下一行开始删除，共删除 dataList.size() 行（即下面的空白行）
                int startDeleteIndex = insertIndex; // 插入结束后的第一行空白行索引
                for (int i = 0; i < dataList.size(); i++) {
                    if (startDeleteIndex < table.getNumberOfRows()) {
                        table.removeRow(startDeleteIndex);
                    }
                }

            }

        }
    }


    /**
     * 检查文本中是否包含指定的字符(此处为“$”)
     *
     * @param text
     * @return
     */
    private static boolean checkText(String text) {
        boolean check = false;
        if (text.contains("$")) {
            check = true;
        }
        return check;
    }

    /**
     * 替换内容
     *
     * @param paragraph
     * @param autEntity
     */
    private static void replaceValueAut(XWPFParagraph paragraph, AutEntity autEntity) {
        List<XWPFRun> runs = paragraph.getRuns();
        for (int i = 0; i < runs.size(); ) {
            XWPFRun run = runs.get(i);
            if (run == null) {
                i++;
                continue;
            }

            String currentText = run.getText(0);
            if (StringUtils.isBlank(currentText)) {
                i++;
                continue;
            }

            // 判断是否开始匹配 ${ 或 $ { 跨 run 的情况
            boolean startsWithDollarBrace = currentText.contains("${");
            boolean potentialSplit = false;

            if (currentText.contains("$") && i + 1 < runs.size()) {
                XWPFRun nextRun = runs.get(i + 1);
                if (nextRun != null) {
                    String nextText = nextRun.getText(0);
                    if (StringUtils.isNotBlank(nextText) && nextText.startsWith("{")) {
                        potentialSplit = true;
                    }
                }
            }

            if (!startsWithDollarBrace && !potentialSplit) {
                i++;
                continue;
            }

            // 开始拼接直到找到右括号 }
            StringBuilder fullText = new StringBuilder(currentText);
            while (i + 1 < runs.size() && !fullText.toString().contains("}")) {
                XWPFRun nextRun = runs.get(i + 1);
                if (nextRun == null) {
                    break;
                }
                String nextText = nextRun.getText(0);
                fullText.append(nextText);
                paragraph.removeRun(i + 1);
            }

            String finalText = fullText.toString();

            // 获取替换值并设置
            Object value = changeValueAut(finalText, autEntity);
            run.setText(value != null ? value.toString() : "", 0);

            i++; // 继续下一个 run
        }
    }

    /**
     * 匹配参数
     *
     * @param value
     * @param autEntity
     * @return
     */
    private static Object changeValueAut(String value, AutEntity autEntity) {
        Object val = "";
        Map<String, Object> autEntityMap = new HashMap<>();
        Arrays.stream(BeanUtils.getPropertyDescriptors(AutEntity.class))
                .filter(pd -> !pd.getName().equals("class"))
                .forEach(pd -> {
                    try {
                        autEntityMap.put(pd.getName(), pd.getReadMethod().invoke(autEntity));
                    } catch (Exception e) {
                        // 处理异常
                    }
                });
        for (Map.Entry<String, Object> textSet : autEntityMap.entrySet()) {
            // 匹配模板与替换值 格式${key}
            String key = textSet.getKey();
            if (value.startsWith("${") && value.endsWith("}")) {
                value = value.substring(2, value.length() - 1);
            }
            if (value.equals(key)) {
                val = textSet.getValue();
                return val;
            }
        }
        return val;
    }

    /**
     * 将Excel文件解析成AUT数据
     *
     * @param file
     * @param sheetAt
     * @param startLine
     * @param endLine
     * @return
     */
    private AutEntity ParsingExcelToAut(MultipartFile file, Integer sheetAt, Integer startLine, Integer endLine) {
        AutEntity autEntity = new AutEntity();
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(file.getInputStream());
        } catch (IOException e) {
            log.error("解析Excel文件失败", e);
            e.printStackTrace();
        } catch (Exception e) {
            log.error("解析Excel文件失败", e);
            e.printStackTrace();
        }

        //获取Sheet
        Sheet sheet = workbook.getSheetAt(sheetAt - 1);
        System.out.println("=============================================================");
        System.out.println("Sheet Name: " + sheet.getSheetName());
        if (!sheet.getSheetName().contains("机组")) {
            System.out.println("该Sheet不是机组Sheet");
            return null;
        }
        // 获取最大行数
        int rownum = sheet.getPhysicalNumberOfRows();
        System.out.println("Row Number: " + rownum);
        System.out.println("=============================================================");
        int heGeShuLiang = 0;
        int buHeGeShuLiang = 0;
        List<List<String>> xuHaoList = new ArrayList<>();
        int number = 1;

        for (int i = startLine - 1; i < endLine; i++) {
            //行
            Row row = sheet.getRow(i);
            // 行数-打印日志用
            int line = i + 1;
            String hanKouBianHao = "";
            if (row != null && !row.toString().isEmpty()) {
                //A 第1个参数（序号"）
                Cell cell1 = row.getCell(0);
                if (cell1 == null) {
                    System.out.println("*****************************************第" + i + "行" + "缺少序号");
                }
                cell1.setCellType(CellType.STRING);
                String xuHao = cell1.getStringCellValue();
                String str1 = String.format("第%s行第%s列————", line, 1);
                System.out.println(str1 + xuHao);

                //E 第5个参数(焊接方法)
                Cell cell5 = row.getCell(4);
                if (cell5 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少焊接方法");
                }
                cell5.setCellType(CellType.STRING);
                String huanjieFangfa = cell5.getStringCellValue();
                String str5 = String.format("第%s行第%s列————", line, 5);
                System.out.println(str5 + huanjieFangfa);
                autEntity.setHanJieFangFa(huanjieFangfa);

                //F 第6个参数(规格尺寸)
                Cell cell6 = row.getCell(5);
                if (cell6 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少规格尺寸");
                }
                cell6.setCellType(CellType.STRING);
                String guiGeChiCun = cell6.getStringCellValue();
                String str6 = String.format("第%s行第%s列————", line, 6);
                System.out.println(str6 + guiGeChiCun);
                autEntity.setGuiGe(guiGeChiCun);

                //G 第7个参数(材质)
                Cell cell7 = row.getCell(6);
                if (cell7 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少材质");
                }
                cell7.setCellType(CellType.STRING);
                String caiZhi = cell7.getStringCellValue();
                String str7 = String.format("第%s行第%s列————", line, 7);
                System.out.println(str7 + caiZhi);
                autEntity.setCaiZhi(caiZhi);

                //H 第8个参数(坡口型式)
                Cell cell8 = row.getCell(7);
                if (cell8 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少坡口形式");
                }
                cell8.setCellType(CellType.STRING);
                String poKouXingShi = cell8.getStringCellValue();
                String str8 = String.format("第%s行第%s列————", line, 8);
                System.out.println(str8 + poKouXingShi);
                autEntity.setPoKouXingShi(poKouXingShi);

                //I 第9个参数(桩号)
                Cell cell9 = row.getCell(8);
                if (cell9 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少桩号");
                }
                cell9.setCellType(CellType.STRING);
                String zhuangHao = cell9.getStringCellValue();
                String str9 = String.format("第%s行第%s列————", line, 9);
                System.out.println(str9 + zhuangHao);
                autEntity.setZhuangHao(zhuangHao);

                //J 第10个参数(焊口编号)
                Cell cell10 = row.getCell(9);
                if (cell10 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少焊口编号");
                }
                cell10.setCellType(CellType.STRING);
                hanKouBianHao = cell10.getStringCellValue();

                String str10 = String.format("第%s行第%s列————", line, 10);
                System.out.println(str10 + hanKouBianHao);

                //N 第14个参数(报告编号)
                Cell cell14 = row.getCell(13);
                if (cell14 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少报告编号");
                }
                cell14.setCellType(CellType.STRING);
                String aauBaogaoBianHao = cell14.getStringCellValue();
                String str14 = String.format("第%s行第%s列————", line, 14);
                System.out.println(str14 + aauBaogaoBianHao);
                autEntity.setBaoGaoBianHao(aauBaogaoBianHao);

                //O 第15个参数(指令时间)
                Cell cell15 = row.getCell(14);
                if (cell15 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少指令时间");
                }
                cell15.setCellType(CellType.STRING);
                String aauZhiLingShiJian = cell15.getStringCellValue();
                String str15 = String.format("第%s行第%s列————", line, 15);
                System.out.println(str15 + aauZhiLingShiJian);
                //格式化
                SimpleDateFormat sdf = new SimpleDateFormat("yyyy年MM月dd日");
                //P 第16个参数(检测时间)
                Cell cell16 = row.getCell(15);
                if (cell16 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少检测时间");
                }
                cell16.setCellType(CellType.NUMERIC);
                Date detectDate = cell16.getDateCellValue();
                if (Objects.nonNull(detectDate)) {
                    String cellValue = sdf.format(detectDate);
                    autEntity.setJianCeRiQi(cellValue);
                }

                //V 第22个参数(评定结果)
                Cell cell22 = row.getCell(21);
                if (cell22 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少评定结果");
                }
                cell22.setCellType(CellType.STRING);
                String aauPingDingJieGuo = cell22.getStringCellValue();
                String str22 = String.format("第%s行第%s列————", line, 22);
                System.out.println(str22 + aauPingDingJieGuo);

                if (aauPingDingJieGuo.equals("合格")) {
                    heGeShuLiang++;
                } else if (aauPingDingJieGuo.equals("不合格")) {
                    buHeGeShuLiang++;
                }

                //Y 第25个参数(备注)
                Cell cell25 = row.getCell(24);
                if (cell25 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少备注");
                }
                cell25.setCellType(CellType.STRING);
                String aauBeiZhu = cell25.getStringCellValue();
                String str25 = String.format("第%s行第%s列————", line, 25);
                System.out.println(str25 + aauBeiZhu);
                //设备型号（V4-200398的话就是Focus LT-RM 32:128。CPP-150032或者PRI-AUT-14160004的话就是CPP-PRI-AUT 32:128）
                //设备型号（TRIASSIC-AUT-V5/18071或者TRIASSIC-AUT-V5/18072，都填TRIASSIC-AUT-V5）
                //检测灵敏度（V4-200398的话就是11。CPP-150032或者PRI-AUT-14160004的话就是8，TRIASSIC-AUT-V5的话就是9）
                if ("V4-200398".equals(aauBeiZhu)) {
                    autEntity.setSheBeiXingHao("Focus LT-RM 32:128");
                    autEntity.setJianCeLingMinDu(8);
                } else if ("CPP-150032".equals(aauBeiZhu) || "PRI-AUT-14160004".equals(aauBeiZhu)) {
                    autEntity.setSheBeiXingHao("CPP-PRI-AUT 32:128");
                    autEntity.setJianCeLingMinDu(11);
                } else if ("TRIASSIC-AUT-V5/18071".equals(aauBeiZhu) || "TRIASSIC-AUT-V5/18072".equals(aauBeiZhu)) {
                    autEntity.setSheBeiXingHao("TRIASSIC-AUT-V5");
                    autEntity.setJianCeLingMinDu(9);
                } else {
                    autEntity.setSheBeiXingHao(aauBeiZhu);
                    autEntity.setJianCeLingMinDu(0);
                }
            } else {
                System.out.println("*****************************************第" + line + "行" + "缺少参数");
            }
            int jianCeShuLiang = endLine - startLine + 1;
            autEntity.setJianCeShuLiang(jianCeShuLiang);
            autEntity.setHeGeShuLiang(heGeShuLiang);
            autEntity.setFanXiuShuLiang(buHeGeShuLiang);
            if (heGeShuLiang == 0) {
                autEntity.setYiCiHeGeLv("0%");
            } else {
                // 使用 double 避免整数除法精度丢失
                double passRate = (double) heGeShuLiang / jianCeShuLiang;
                // 格式化输出百分比（保留两位小数）
                String passRateStr = String.format("%.2f", passRate * 100);
                autEntity.setYiCiHeGeLv(passRateStr);
            }

            List<String> xuHaoDataList = new ArrayList<>();
            xuHaoDataList.add(String.valueOf(number));
            xuHaoDataList.add(hanKouBianHao);
            DataFormatter dataFormatter = new DataFormatter();
            for (int colIndex = 16; colIndex <= 21; colIndex++) {
                Cell cell = row.getCell(colIndex);
                if (cell != null) {
                    String str = dataFormatter.formatCellValue(cell);
                    xuHaoDataList.add(str);
                } else {
                    xuHaoDataList.add(""); // Add empty string if cell is null
                }
            }
            xuHaoDataList.add("/");
            number = number + 1;
            xuHaoList.add(xuHaoDataList);
        }
        List<String> kongDataList = new ArrayList<>();
        kongDataList.add("");
        kongDataList.add("以下空白");
        kongDataList.add("");
        kongDataList.add("");
        kongDataList.add("");
        kongDataList.add("");
        kongDataList.add("");
        kongDataList.add("");
        kongDataList.add("");
        xuHaoList.add(kongDataList);
        autEntity.setXuHao(xuHaoList);
        return autEntity;
    }


    /**
     * 获取单元格文本内容（合并 run 的情况）
     */
    private static String getCellText(XWPFTableCell cell) {
        StringBuilder sb = new StringBuilder();
        for (XWPFParagraph p : cell.getParagraphs()) {
            for (XWPFRun r : p.getRuns()) {
                sb.append(r.getText(0));
            }
        }
        return sb.toString().trim();
    }


    /**
     * 合并表格中某一行的指定列范围的单元格（横向合并）
     *
     * @param row      要合并的行
     * @param fromCell 起始列索引（从0开始）
     * @param toCell   结束列索引（包含）
     */
    public static void mergeCellsHorizontal(XWPFTableRow row, int fromCell, int toCell) {
        if (fromCell < 0 || toCell < 0 || fromCell > toCell) {
            throw new IllegalArgumentException("Invalid cell range");
        }
        XWPFTableCell cell = row.getCell(fromCell);
        if (cell == null) return;
        CTTc ctTc = cell.getCTTc();
        // 获取或创建单元格属性
        CTTcPr tcPr = ctTc.isSetTcPr() ? ctTc.getTcPr() : ctTc.addNewTcPr();
        // 设置 gridSpan 属性，表示合并的列数
        tcPr.addNewGridSpan().setVal(BigInteger.valueOf(toCell - fromCell + 1));
        // 清空被合并的其他单元格内容
        for (int i = fromCell + 1; i <= toCell; i++) {
            XWPFTableCell nextCell = row.getCell(i);
            if (nextCell != null) {
                CTTc nextTc = nextCell.getCTTc();
                // 可选：清除原属性
                nextTc.setTcPr(null);
                // 垂直居中
                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            }
        }
    }


////////////以下为处理DR逻辑


    /**
     * 将Excel文件解析成DR数据
     *
     * @param file
     * @param sheetAt
     * @param startLine
     * @param endLine
     * @return
     */
    private DrEntity ParsingExcelToDr(MultipartFile file, Integer sheetAt, Integer startLine, Integer endLine) {
        DrEntity drEntity = new DrEntity();
        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(file.getInputStream());
        } catch (IOException e) {
            log.error("解析Excel文件失败", e);
            e.printStackTrace();
        } catch (Exception e) {
            log.error("解析Excel文件失败", e);
            e.printStackTrace();
        }

        //获取Sheet
        Sheet sheet = workbook.getSheetAt(sheetAt - 1);
        System.out.println("=============================================================");
        System.out.println("Sheet Name: " + sheet.getSheetName());
        if (!sheet.getSheetName().contains("机组")) {
            System.out.println("该Sheet不是机组Sheet");
            return null;
        }
        // 获取最大行数
        int rownum = sheet.getPhysicalNumberOfRows();
        System.out.println("Row Number: " + rownum);
        System.out.println("=============================================================");
        int heGeShuLiang = 0;
        int buHeGeShuLiang = 0;
        List<List<String>> xuHaoList = new ArrayList<>();
        int number = 1;

        for (int i = startLine - 1; i < endLine; i++) {
            //行
            Row row = sheet.getRow(i);
            // 行数-打印日志用
            int line = i + 1;
            String hanKouBianHao = "";
            String guiGe = "φ";
            String xianXingXiangZhiJiZhiShu = "";
            String shuangXianXingXiangZhiJiZhiShu = "";
            String queQian = "";
            String panDingJieGuo = "";
            if (row != null && !row.toString().isEmpty()) {
                //E 第5个参数(焊接方法)
                Cell cell5 = row.getCell(4);
                if (cell5 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少焊接方法");
                }
                cell5.setCellType(CellType.STRING);
                String huanjieFangfa = cell5.getStringCellValue();
                String str5 = String.format("第%s行第%s列————", line, 5);
                System.out.println(str5 + huanjieFangfa);
                drEntity.setHanJieFangFa(huanjieFangfa);

                //F 第6个参数(管材规格)
                Cell cell6 = row.getCell(5);
                if (cell6 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少规格尺寸");
                }
                cell6.setCellType(CellType.STRING);
                guiGe = cell6.getStringCellValue();
                String str6 = String.format("第%s行第%s列————", line, 6);
                System.out.println(str6 + guiGe);
                drEntity.setGuanCaiGuiGe(guiGe);
                if ("φ1219×18.4".equals(guiGe)) {
                    drEntity.setGuanDianYa("260");
                } else if ("φ1219×22".equals(guiGe)) {
                    drEntity.setGuanDianYa("270");
                }

                //H 第8个参数(坡口型式)
                Cell cell8 = row.getCell(7);
                if (cell8 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少坡口形式");
                }
                cell8.setCellType(CellType.STRING);
                String poKouXingShi = cell8.getStringCellValue();
                String str8 = String.format("第%s行第%s列————", line, 8);
                System.out.println(str8 + poKouXingShi);
                drEntity.setPoKouXingShi(poKouXingShi);

                //I 第9个参数(桩号)
                Cell cell9 = row.getCell(8);
                if (cell9 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少桩号");
                }
                cell9.setCellType(CellType.STRING);
                String zhuangHao = cell9.getStringCellValue();
                String str9 = String.format("第%s行第%s列————", line, 9);
                System.out.println(str9 + zhuangHao);
                drEntity.setZhuangHao(zhuangHao);

                //J 第8个参数(焊缝编号)
                Cell cell10 = row.getCell(9);
                if (cell10 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少焊缝编号");
                }
                cell10.setCellType(CellType.STRING);
                hanKouBianHao = cell10.getStringCellValue();
                String str10 = String.format("第%s行第%s列————", line, 10);
                System.out.println(str10 + hanKouBianHao);

                //AE 第31个参数(指令编号)
                Cell cell31 = row.getCell(30);
                if (cell31 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少指令编号");
                }
                cell31.setCellType(CellType.STRING);
                String zhiLingBianHao = cell31.getStringCellValue();
                String str31 = String.format("第%s行第%s列————", line, 31);
                System.out.println(str31 + zhiLingBianHao);
                drEntity.setZhiLingBianHao(zhiLingBianHao);

                //AF 第32个参数(报告编号)
                Cell cell32 = row.getCell(31);
                if (cell32 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少报告编号");
                }
                cell32.setCellType(CellType.STRING);
                String baogaoBianHao = cell32.getStringCellValue();
                String str32 = String.format("第%s行第%s列————", line, 32);
                System.out.println(str32 + baogaoBianHao);
                drEntity.setBaoGaoBianHao(baogaoBianHao);

                SimpleDateFormat sdf = new SimpleDateFormat("yyyy年MM月dd日");
                //AH 第34个参数(检测日期)
                Cell cell34 = row.getCell(33);
                if (cell34 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少检测日期");
                }
                cell34.setCellType(CellType.NUMERIC);
                Date jianCeRiQi = cell34.getDateCellValue();
                if (Objects.nonNull(jianCeRiQi)) {
                    String cellValue = sdf.format(jianCeRiQi);
                    drEntity.setJianCeRiQi(cellValue);
                }
                String str34 = String.format("第%s行第%s列————", line, 34);
                System.out.println(str34 + jianCeRiQi);

                //AK 第37个参数(检测日期)
                Cell cell37 = row.getCell(36);
                if (cell37 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少检测日期");
                }
                cell37.setCellType(CellType.STRING);
                String huiDuZhiFanWei = cell37.getStringCellValue();
                String str37 = String.format("第%s行第%s列————", line, 37);
                System.out.println(str37 + huiDuZhiFanWei);
                drEntity.setHuiDuZhiFanWei(huiDuZhiFanWei);

                //AN 第40个参数(线型像质计指数)
                Cell cell40 = row.getCell(39);
                if (cell40 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少线型像质计指数");
                }
                cell40.setCellType(CellType.STRING);
                xianXingXiangZhiJiZhiShu = cell40.getStringCellValue();
                String str40 = String.format("第%s行第%s列————", line, 40);
                System.out.println(str40 + xianXingXiangZhiJiZhiShu);

                //AO 第41个参数(双线型像质计指数)
                Cell cell41 = row.getCell(40);
                if (cell41 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少线型像质计指数");
                }
                cell41.setCellType(CellType.STRING);
                shuangXianXingXiangZhiJiZhiShu = cell41.getStringCellValue();
                String str41 = String.format("第%s行第%s列————", line, 41);
                System.out.println(str41 + shuangXianXingXiangZhiJiZhiShu);

                //AX 第50个参数(探测器型号)
                Cell cell50 = row.getCell(49);
                if (cell50 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少探测器型号");
                }
                cell50.setCellType(CellType.STRING);
                String tanCeQiXingHao = cell50.getStringCellValue();
                String str50 = String.format("第%s行第%s列————", line, 50);
                System.out.println(str50 + tanCeQiXingHao);

                if ("DY-ZR-220413701".equals(tanCeQiXingHao)) {
                    drEntity.setTanCeQiGuiGe("285×250mm");
                    drEntity.setGongYiKaBianHao("HL/CQDS2-DRGYK-003-2025");
                    drEntity.setTanCeQiXingHao("RAPIXX 3NDT WIFI");
                    drEntity.setXiangSuChiCun("139");
                    drEntity.setYiCiTouZhaoChangDu("226");
                    if ("φ1219×18.4".equals(guiGe)) {
                        drEntity.setZongBaoGuangShiJian("8.55");
                    } else if ("φ1219×22".equals(guiGe)) {
                        drEntity.setZongBaoGuangShiJian("9");
                    } else if ("φ1219×27.5".equals(guiGe)) {
                        drEntity.setZongBaoGuangShiJian("9.2");
                    }
                } else if ("CPP1784022".equals(tanCeQiXingHao)) {
                    drEntity.setTanCeQiGuiGe("274×184mm");
                    drEntity.setGongYiKaBianHao("HL/CQDS2-DRGYK-004-2025");
                    drEntity.setTanCeQiXingHao("RAPIXX 2NDT WIFI");
                    drEntity.setXiangSuChiCun("125");
                    drEntity.setYiCiTouZhaoChangDu("119");
                    if ("φ1219×18.4".equals(guiGe) || "φ1219×22".equals(guiGe)) {
                        drEntity.setZongBaoGuangShiJian("9.2");
                    }
                }

                //AU 第49个参数(评定结果)
                Cell cell47 = row.getCell(46);
                if (cell47 == null) {
                    System.out.println("*****************************************第" + line + "行" + "评定结果");
                }
                cell47.setCellType(CellType.STRING);
                panDingJieGuo = cell47.getStringCellValue();
                String str47 = String.format("第%s行第%s列————", line, 49);
                System.out.println(str47 + panDingJieGuo);

                //AV 第48个参数(评定结果)
                Cell cell48 = row.getCell(47);
                if (cell48 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少评定结果");
                }
                cell48.setCellType(CellType.STRING);
                String pingDingJieGuo = cell48.getStringCellValue();
                String str48 = String.format("第%s行第%s列————", line, 48);
                System.out.println(str48 + pingDingJieGuo);
                if (pingDingJieGuo.equals("合格")) {
                    heGeShuLiang++;
                } else if (pingDingJieGuo.equals("不合格")) {
                    buHeGeShuLiang++;
                }

                //AW 第49个参数(缺欠位置、性质及长度)
                Cell cell49 = row.getCell(48);
                if (cell49 == null) {
                    System.out.println("*****************************************第" + line + "行" + "缺少缺欠位置、性质及长度");
                }
                cell49.setCellType(CellType.STRING);
                queQian = cell49.getStringCellValue();
                String str49 = String.format("第%s行第%s列————", line, 49);
                System.out.println(str49 + queQian);


            } else {
                System.out.println("*****************************************第" + line + "行" + "缺少参数");
            }
            int jianCeShuLiang = endLine - startLine + 1;
            drEntity.setJianCeShuLiang(jianCeShuLiang);
            drEntity.setHeGeShuLiang(heGeShuLiang);
            drEntity.setFanXiuShuLiang(buHeGeShuLiang);
            if (heGeShuLiang == 0) {
                drEntity.setYiCiHeGeLv("0%");
            } else {
                // 使用 double 避免整数除法精度丢失
                double passRate = (double) heGeShuLiang / jianCeShuLiang;
                // 格式化输出百分比（保留两位小数）
                String passRateStr = String.format("%.2f", passRate * 100);
                drEntity.setYiCiHeGeLv(passRateStr);
            }

            List<String> xuHaoDataList = new ArrayList<>();
            xuHaoDataList.add(String.valueOf(number));
            xuHaoDataList.add(hanKouBianHao);
            xuHaoDataList.add(guiGe.replace("φ", ""));
            xuHaoDataList.add(drEntity.getYiCiTouZhaoChangDu());
            xuHaoDataList.add(xianXingXiangZhiJiZhiShu);
            xuHaoDataList.add(shuangXianXingXiangZhiJiZhiShu);
            xuHaoDataList.add(queQian);
            xuHaoDataList.add(panDingJieGuo);
            xuHaoDataList.add("/");
            number = number + 1;
            xuHaoList.add(xuHaoDataList);
        }
        List<String> kongDataList = new ArrayList<>();
        kongDataList.add("");
        kongDataList.add("以下空白");
        kongDataList.add("");
        kongDataList.add("");
        kongDataList.add("");
        kongDataList.add("");
        kongDataList.add("");
        kongDataList.add("");
        kongDataList.add("");
        xuHaoList.add(kongDataList);
        drEntity.setXuHao(xuHaoList);
        return drEntity;
    }

    /**
     * 替换表格内的文字
     *
     * @param document
     * @param drEntity
     */
    private static void changeTableTextDr(XWPFDocument document, DrEntity drEntity) {
        Iterator<XWPFTable> tableIt = document.getTablesIterator();
        while (tableIt.hasNext()) {
            XWPFTable table = tableIt.next();
            if (!checkText(table.getText())) {
                continue;
            }

            List<Integer> replaceIndexes = new ArrayList<>();
            List<XWPFTableRow> rows = table.getRows();

            for (int i = 0; i < rows.size(); i++) {
                XWPFTableRow row = rows.get(i);
                List<XWPFTableCell> cells = row.getTableCells();
                if (cells == null || cells.isEmpty()) continue;

                for (XWPFTableCell cell : cells) {
                    String text = getCellText(cell).trim();
                    if ("${xuHao}".equals(text)) {
                        replaceIndexes.add(i);
                        break;
                    }

                    // 缓存单元格文本，避免重复调用 getText()
                    String cellText = cell.getText();
                    if (checkText(cellText)) {
                        List<XWPFParagraph> paragraphs = cell.getParagraphs();
                        if (paragraphs != null && !paragraphs.isEmpty()) {
                            for (XWPFParagraph paragraph : paragraphs) {
                                replaceValueDr(paragraph, drEntity);
                            }
                        }
                    }
                }
            }

            // 倒序处理，防止插入影响后续索引
            Collections.sort(replaceIndexes, Collections.reverseOrder());

            for (Integer rowIndex : replaceIndexes) {
                List<List<String>> dataList = drEntity.getXuHao();
                if (dataList == null || dataList.isEmpty()) continue;

                int insertIndex = rowIndex + 1; // 在占位符行的下一行插入

                for (List<String> rowData : dataList) {
                    XWPFTableRow newRow = table.insertNewTableRow(insertIndex++);
                    // 设置行高0.85厘米
                    newRow.setHeight(482);

                    for (String cellValue : rowData) {
                        XWPFTableCell cell = newRow.addNewTableCell();
                        cell.removeParagraph(0); // 删除默认段落
                        XWPFParagraph para = cell.addParagraph();
                        para.setAlignment(ParagraphAlignment.CENTER);
                        para.setSpacingBefore(0);// 段前间距
                        para.setSpacingAfter(0);// 段后间距
                        XWPFRun run = para.createRun();
                        run.setText(cellValue);
                        run.setFontSize(10.5); // 设置为5号字体
                        run.setFontFamily("宋体");
                        // 设置垂直居中
                        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

                        //设置行间距为12磅（1磅=20 twip，12磅=240 twip）
                        CTP ctp = para.getCTP();
                        CTPPr ppr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
                        CTSpacing spacing = ppr.isSetSpacing() ? ppr.getSpacing() : ppr.addNewSpacing();
                        spacing.setLineRule(STLineSpacingRule.EXACT);
                        spacing.setLine(BigInteger.valueOf(240));
                    }
                    // 合并单元格逻辑...
                    if (rowData.size() > 1) {
                        mergeCellsHorizontal(newRow, 1, 3);
                        mergeCellsHorizontal(newRow, 6, 8);
                    }
                }
                // 可选：插入完成后删除原始占位符行
                table.removeRow(rowIndex);
                // 然后从插入内容的下一行开始删除，共删除 dataList.size() 行（即下面的空白行）
                int startDeleteIndex = insertIndex; // 插入结束后的第一行空白行索引
                for (int i = 0; i < dataList.size(); i++) {
                    if (startDeleteIndex < table.getNumberOfRows()) {
                        table.removeRow(startDeleteIndex);
                    }
                }
            }
        }
    }

    /**
     * 替换内容
     *
     * @param paragraph
     * @param drEntity
     */
    private static void replaceValueDr(XWPFParagraph paragraph, DrEntity drEntity) {
        List<XWPFRun> runs = paragraph.getRuns();
        for (int i = 0; i < runs.size(); ) {
            XWPFRun run = runs.get(i);
            if (run == null) {
                i++;
                continue;
            }

            String currentText = run.getText(0);
            if (StringUtils.isBlank(currentText)) {
                i++;
                continue;
            }

            // 判断是否开始匹配 ${ 或 $ { 跨 run 的情况
            boolean startsWithDollarBrace = currentText.contains("${");
            boolean potentialSplit = false;

            if (currentText.contains("$") && i + 1 < runs.size()) {
                XWPFRun nextRun = runs.get(i + 1);
                if (nextRun != null) {
                    String nextText = nextRun.getText(0);
                    if (StringUtils.isNotBlank(nextText) && nextText.startsWith("{")) {
                        potentialSplit = true;
                    }
                }
            }

            if (!startsWithDollarBrace && !potentialSplit) {
                i++;
                continue;
            }

            // 开始拼接直到找到右括号 }
            StringBuilder fullText = new StringBuilder(currentText);
            while (i + 1 < runs.size() && !fullText.toString().contains("}")) {
                XWPFRun nextRun = runs.get(i + 1);
                if (nextRun == null) {
                    break;
                }
                String nextText = nextRun.getText(0);
                fullText.append(nextText);
                paragraph.removeRun(i + 1);
            }

            String finalText = fullText.toString();

            // 获取替换值并设置
            Object value = changeValueDr(finalText, drEntity);
            run.setText(value != null ? value.toString() : "", 0);

            i++; // 继续下一个 run
        }
    }


    /**
     * 匹配参数
     *
     * @param value
     * @param drEntity
     * @return
     */
    private static Object changeValueDr(String value, DrEntity drEntity) {
        Object val = "";
        Map<String, Object> drEntityMap = new HashMap<>();
        Arrays.stream(BeanUtils.getPropertyDescriptors(DrEntity.class))
                .filter(pd -> !pd.getName().equals("class"))
                .forEach(pd -> {
                    try {
                        drEntityMap.put(pd.getName(), pd.getReadMethod().invoke(drEntity));
                    } catch (Exception e) {
                        // 处理异常
                    }
                });
        for (Map.Entry<String, Object> textSet : drEntityMap.entrySet()) {
            // 匹配模板与替换值 格式${key}
            String key = textSet.getKey();
            if (value.startsWith("${") && value.endsWith("}")) {
                value = value.substring(2, value.length() - 1);
            }
            if (value.equals(key)) {
                val = textSet.getValue();
                return val;
            }
        }
        return val;
    }


////////////以下为处理PA逻辑
//
//    /**
//     * 将Excel文件解析成PA数据
//     *
//     * @param file
//     * @param sheetAt
//     * @param startLine
//     * @param endLine
//     * @return
//     */
//    private PaEntity ParsingExcelToPa(MultipartFile file, Integer sheetAt, Integer startLine, Integer endLine) {
//        PaEntity paEntity = new PaEntity();
//        Workbook workbook = null;
//        try {
//            workbook = WorkbookFactory.create(file.getInputStream());
//        } catch (IOException e) {
//            log.error("解析Excel文件失败", e);
//            e.printStackTrace();
//        } catch (Exception e) {
//            log.error("解析Excel文件失败", e);
//            e.printStackTrace();
//        }
//
//        //获取Sheet
//        Sheet sheet = workbook.getSheetAt(sheetAt - 1);
//        System.out.println("=============================================================");
//        System.out.println("Sheet Name: " + sheet.getSheetName());
//        if (!sheet.getSheetName().contains("机组")) {
//            System.out.println("该Sheet不是机组Sheet");
//            return null;
//        }
//        // 获取最大行数
//        int rownum = sheet.getPhysicalNumberOfRows();
//        System.out.println("Row Number: " + rownum);
//        System.out.println("=============================================================");
//        int heGeShuLiang = 0;
//        int buHeGeShuLiang = 0;
//        List<List<String>> xuHaoList = new ArrayList<>();
//        int number = 1;
//
//        for (int i = startLine - 1; i < endLine; i++) {
//            //行
//            Row row = sheet.getRow(i);
//            // 行数-打印日志用
//            int line = i + 1;
//            String hanKouBianHao = "";
//            String guiGe = "φ";
//            String xianXingXiangZhiJiZhiShu = "";
//            String shuangXianXingXiangZhiJiZhiShu = "";
//            String queQian = "";
//            String panDingJieGuo = "";
//            if (row != null && !row.toString().isEmpty()) {
//                //E 第5个参数(焊接方法)
//                Cell cell5 = row.getCell(4);
//                if (cell5 == null) {
//                    System.out.println("*****************************************第" + line + "行" + "缺少焊接方法");
//                }
//                cell5.setCellType(CellType.STRING);
//                String huanjieFangfa = cell5.getStringCellValue();
//                String str5 = String.format("第%s行第%s列————", line, 5);
//                System.out.println(str5 + huanjieFangfa);
//                paEntity.setHanJieFangFa(huanjieFangfa);
//
//                //F 第6个参数(管材规格)
//                Cell cell6 = row.getCell(5);
//                if (cell6 == null) {
//                    System.out.println("*****************************************第" + line + "行" + "缺少规格尺寸");
//                }
//                cell6.setCellType(CellType.STRING);
//                guiGe = cell6.getStringCellValue();
//                String str6 = String.format("第%s行第%s列————", line, 6);
//                System.out.println(str6 + guiGe);
//                paEntity.setGuanCaiGuiGe(guiGe);
//
//                //H 第8个参数(坡口型式)
//                Cell cell8 = row.getCell(7);
//                if (cell8 == null) {
//                    System.out.println("*****************************************第" + line + "行" + "缺少坡口形式");
//                }
//                cell8.setCellType(CellType.STRING);
//                String poKouXingShi = cell8.getStringCellValue();
//                String str8 = String.format("第%s行第%s列————", line, 8);
//                System.out.println(str8 + poKouXingShi);
//                paEntity.setPoKouXingShi(poKouXingShi);
//
//                //I 第9个参数(桩号)
//                Cell cell9 = row.getCell(8);
//                if (cell9 == null) {
//                    System.out.println("*****************************************第" + line + "行" + "缺少桩号");
//                }
//                cell9.setCellType(CellType.STRING);
//                String zhuangHao = cell9.getStringCellValue();
//                String str9 = String.format("第%s行第%s列————", line, 9);
//                System.out.println(str9 + zhuangHao);
//                paEntity.setZhuangHao(zhuangHao);
//
//                //J 第8个参数(焊缝编号)
//                Cell cell10 = row.getCell(9);
//                if (cell10 == null) {
//                    System.out.println("*****************************************第" + line + "行" + "缺少焊缝编号");
//                }
//                cell10.setCellType(CellType.STRING);
//                hanKouBianHao = cell10.getStringCellValue();
//                String str10 = String.format("第%s行第%s列————", line, 10);
//                System.out.println(str10 + hanKouBianHao);
//
//                //AE 第31个参数(指令编号)
//                Cell cell31 = row.getCell(30);
//                if (cell31 == null) {
//                    System.out.println("*****************************************第" + line + "行" + "缺少指令编号");
//                }
//                cell31.setCellType(CellType.STRING);
//                String zhiLingBianHao = cell31.getStringCellValue();
//                String str31 = String.format("第%s行第%s列————", line, 31);
//                System.out.println(str31 + zhiLingBianHao);
//                paEntity.setZhiLingBianHao(zhiLingBianHao);
//
//                //AF 第32个参数(报告编号)
//                Cell cell32 = row.getCell(31);
//                if (cell32 == null) {
//                    System.out.println("*****************************************第" + line + "行" + "缺少报告编号");
//                }
//                cell32.setCellType(CellType.STRING);
//                String baogaoBianHao = cell32.getStringCellValue();
//                String str32 = String.format("第%s行第%s列————", line, 32);
//                System.out.println(str32 + baogaoBianHao);
//                paEntity.setBaoGaoBianHao(baogaoBianHao);
//
//                SimpleDateFormat sdf = new SimpleDateFormat("yyyy年MM月dd日");
//                //AH 第34个参数(检测日期)
//                Cell cell34 = row.getCell(33);
//                if (cell34 == null) {
//                    System.out.println("*****************************************第" + line + "行" + "缺少检测日期");
//                }
//                cell34.setCellType(CellType.NUMERIC);
//                Date jianCeRiQi = cell34.getDateCellValue();
//                if (Objects.nonNull(jianCeRiQi)) {
//                    String cellValue = sdf.format(jianCeRiQi);
//                    paEntity.setJianCeRiQi(cellValue);
//                }
//                String str34 = String.format("第%s行第%s列————", line, 34);
//                System.out.println(str34 + jianCeRiQi);
//
//                //AK 第37个参数(检测日期)
//                Cell cell37 = row.getCell(36);
//                if (cell37 == null) {
//                    System.out.println("*****************************************第" + line + "行" + "缺少检测日期");
//                }
//                cell37.setCellType(CellType.STRING);
//                String huiDuZhiFanWei = cell37.getStringCellValue();
//                String str37 = String.format("第%s行第%s列————", line, 37);
//                System.out.println(str37 + huiDuZhiFanWei);
//                paEntity.setHuiDuZhiFanWei(huiDuZhiFanWei);
//
//                //AN 第40个参数(线型像质计指数)
//                Cell cell40 = row.getCell(39);
//                if (cell40 == null) {
//                    System.out.println("*****************************************第" + line + "行" + "缺少线型像质计指数");
//                }
//                cell40.setCellType(CellType.STRING);
//                xianXingXiangZhiJiZhiShu = cell40.getStringCellValue();
//                String str40 = String.format("第%s行第%s列————", line, 40);
//                System.out.println(str40 + xianXingXiangZhiJiZhiShu);
//
//                //AO 第41个参数(双线型像质计指数)
//                Cell cell41 = row.getCell(40);
//                if (cell41 == null) {
//                    System.out.println("*****************************************第" + line + "行" + "缺少线型像质计指数");
//                }
//                cell41.setCellType(CellType.STRING);
//                shuangXianXingXiangZhiJiZhiShu = cell41.getStringCellValue();
//                String str41 = String.format("第%s行第%s列————", line, 41);
//                System.out.println(str41 + shuangXianXingXiangZhiJiZhiShu);
//
//                //AX 第50个参数(探测器型号)
//                Cell cell50 = row.getCell(49);
//                if (cell50 == null) {
//                    System.out.println("*****************************************第" + line + "行" + "缺少探测器型号");
//                }
//                cell50.setCellType(CellType.STRING);
//                String tanCeQiXingHao = cell50.getStringCellValue();
//                String str50 = String.format("第%s行第%s列————", line, 50);
//                System.out.println(str50 + tanCeQiXingHao);
//
//                if ("DY-ZR-220413701".equals(tanCeQiXingHao)) {
//                    drEntity.setGongYiKaBianHao("HL/CQDS2-DRGYK-003-2025");
//                    drEntity.setTanCeQiXingHao("RAPIXX 3NDT WIFI");
//                    drEntity.setXiangSuChiCun("139");
//                    drEntity.setYiCiTouZhaoChangDu("226");
//                    if ("φ1219×18.4".equals(guiGe)) {
//                        drEntity.setZongBaoGuangShiJian("8.55");
//                    } else if ("φ1219×22".equals(guiGe)) {
//                        drEntity.setZongBaoGuangShiJian("9");
//                    } else if ("φ1219×27.5".equals(guiGe)) {
//                        drEntity.setZongBaoGuangShiJian("9.2");
//                    }
//                } else if ("CPP1784022".equals(tanCeQiXingHao)) {
//                    drEntity.setGongYiKaBianHao("HL/CQDS2-DRGYK-004-2025");
//                    drEntity.setTanCeQiXingHao("RAPIXX 2NDT WIFI");
//                    drEntity.setXiangSuChiCun("125");
//                    drEntity.setYiCiTouZhaoChangDu("119");
//                    if ("φ1219×18.4".equals(guiGe) || "φ1219×22".equals(guiGe)) {
//                        drEntity.setZongBaoGuangShiJian("9.2");
//                    }
//                }
//
//                //AU 第49个参数(评定结果)
//                Cell cell47 = row.getCell(46);
//                if (cell47 == null) {
//                    System.out.println("*****************************************第" + line + "行" + "评定结果");
//                }
//                cell47.setCellType(CellType.STRING);
//                panDingJieGuo = cell47.getStringCellValue();
//                String str47 = String.format("第%s行第%s列————", line, 49);
//                System.out.println(str47 + panDingJieGuo);
//
//                //AV 第48个参数(评定结果)
//                Cell cell48 = row.getCell(47);
//                if (cell48 == null) {
//                    System.out.println("*****************************************第" + line + "行" + "缺少评定结果");
//                }
//                cell48.setCellType(CellType.STRING);
//                String pingDingJieGuo = cell48.getStringCellValue();
//                String str48 = String.format("第%s行第%s列————", line, 48);
//                System.out.println(str48 + pingDingJieGuo);
//                if (pingDingJieGuo.equals("合格")) {
//                    heGeShuLiang++;
//                } else if (pingDingJieGuo.equals("不合格")) {
//                    buHeGeShuLiang++;
//                }
//
//                //AW 第49个参数(缺欠位置、性质及长度)
//                Cell cell49 = row.getCell(48);
//                if (cell49 == null) {
//                    System.out.println("*****************************************第" + line + "行" + "缺少缺欠位置、性质及长度");
//                }
//                cell49.setCellType(CellType.STRING);
//                queQian = cell49.getStringCellValue();
//                String str49 = String.format("第%s行第%s列————", line, 49);
//                System.out.println(str49 + queQian);
//
//
//            } else {
//                System.out.println("*****************************************第" + line + "行" + "缺少参数");
//            }
//            int jianCeShuLiang = endLine - startLine + 1;
//            drEntity.setJianCeShuLiang(jianCeShuLiang);
//            drEntity.setHeGeShuLiang(heGeShuLiang);
//            drEntity.setFanXiuShuLiang(buHeGeShuLiang);
//            if (heGeShuLiang == 0) {
//                drEntity.setYiCiHeGeLv("0%");
//            } else {
//                // 使用 double 避免整数除法精度丢失
//                double passRate = (double) heGeShuLiang / jianCeShuLiang;
//                // 格式化输出百分比（保留两位小数）
//                String passRateStr = String.format("%.2f", passRate * 100);
//                drEntity.setYiCiHeGeLv(passRateStr);
//            }
//
//            List<String> xuHaoDataList = new ArrayList<>();
//            xuHaoDataList.add(String.valueOf(number));
//            xuHaoDataList.add(hanKouBianHao);
//            xuHaoDataList.add(guiGe.replace("φ", ""));
//            xuHaoDataList.add(drEntity.getYiCiTouZhaoChangDu());
//            xuHaoDataList.add(xianXingXiangZhiJiZhiShu);
//            xuHaoDataList.add(shuangXianXingXiangZhiJiZhiShu);
//            xuHaoDataList.add(queQian);
//            xuHaoDataList.add(panDingJieGuo);
//            xuHaoDataList.add("/");
//            number = number + 1;
//            xuHaoList.add(xuHaoDataList);
//        }
//        List<String> kongDataList = new ArrayList<>();
//        kongDataList.add("");
//        kongDataList.add("以下空白");
//        kongDataList.add("");
//        kongDataList.add("");
//        kongDataList.add("");
//        kongDataList.add("");
//        kongDataList.add("");
//        kongDataList.add("");
//        kongDataList.add("");
//        xuHaoList.add(kongDataList);
//        paEntity.setXuHao(xuHaoList);
//        return paEntity;
//    }
//
//    /**
//     * 替换表格内的文字
//     *
//     * @param document
//     * @param paEntity
//     */
//    private static void changeTableTextPa(XWPFDocument document, PaEntity paEntity) {
//        Iterator<XWPFTable> tableIt = document.getTablesIterator();
//        while (tableIt.hasNext()) {
//            XWPFTable table = tableIt.next();
//            if (!checkText(table.getText())) {
//                continue;
//            }
//
//            List<Integer> replaceIndexes = new ArrayList<>();
//            List<XWPFTableRow> rows = table.getRows();
//
//            for (int i = 0; i < rows.size(); i++) {
//                XWPFTableRow row = rows.get(i);
//                List<XWPFTableCell> cells = row.getTableCells();
//                if (cells == null || cells.isEmpty()) continue;
//
//                for (XWPFTableCell cell : cells) {
//                    String text = getCellText(cell).trim();
//                    if ("${xuHao}".equals(text)) {
//                        replaceIndexes.add(i);
//                        break;
//                    }
//
//                    // 缓存单元格文本，避免重复调用 getText()
//                    String cellText = cell.getText();
//                    if (checkText(cellText)) {
//                        List<XWPFParagraph> paragraphs = cell.getParagraphs();
//                        if (paragraphs != null && !paragraphs.isEmpty()) {
//                            for (XWPFParagraph paragraph : paragraphs) {
//                                replaceValuePa(paragraph, paEntity);
//                            }
//                        }
//                    }
//                }
//            }
//
//            // 倒序处理，防止插入影响后续索引
//            Collections.sort(replaceIndexes, Collections.reverseOrder());
//
//            for (Integer rowIndex : replaceIndexes) {
//                List<List<String>> dataList = paEntity.getXuHao();
//                if (dataList == null || dataList.isEmpty()) continue;
//                int insertIndex = rowIndex + 1; // 在占位符行的下一行插入
//
//                for (List<String> rowData : dataList) {
//                    XWPFTableRow newRow = table.insertNewTableRow(insertIndex++);
//                    // 设置行高0.85厘米
//                    newRow.setHeight(482);
//
//                    for (String cellValue : rowData) {
//                        XWPFTableCell cell = newRow.addNewTableCell();
//                        cell.removeParagraph(0); // 删除默认段落
//                        XWPFParagraph para = cell.addParagraph();
//                        para.setAlignment(ParagraphAlignment.CENTER);
//                        para.setSpacingBefore(0);// 段前间距
//                        para.setSpacingAfter(0);// 段后间距
//                        XWPFRun run = para.createRun();
//                        run.setText(cellValue);
//                        run.setFontSize(10.5); // 设置为5号字体
//                        run.setFontFamily("宋体");
//                        // 设置垂直居中
//                        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//
//                        //设置行间距为12磅（1磅=20 twip，12磅=240 twip）
//                        CTP ctp = para.getCTP();
//                        CTPPr ppr = ctp.isSetPPr() ? ctp.getPPr() : ctp.addNewPPr();
//                        CTSpacing spacing = ppr.isSetSpacing() ? ppr.getSpacing() : ppr.addNewSpacing();
//                        spacing.setLineRule(STLineSpacingRule.EXACT);
//                        spacing.setLine(BigInteger.valueOf(240));
//                    }
//                    // 合并单元格逻辑...
//                    if (rowData.size() > 1) {
////                        mergeCellsHorizontal(newRow, 1, 3);
////                        mergeCellsHorizontal(newRow, 6, 8);
//                    }
//                }
//                // 可选：插入完成后删除原始占位符行
//                table.removeRow(rowIndex);
//                // 然后从插入内容的下一行开始删除，共删除 dataList.size() 行（即下面的空白行）
//                int startDeleteIndex = insertIndex; // 插入结束后的第一行空白行索引
//                for (int i = 0; i < dataList.size(); i++) {
//                    if (startDeleteIndex < table.getNumberOfRows()) {
//                        table.removeRow(startDeleteIndex);
//                    }
//                }
//            }
//        }
//    }
//
//    /**
//     * 替换内容
//     *
//     * @param paragraph
//     * @param paEntity
//     */
//    private static void replaceValuePa(XWPFParagraph paragraph, PaEntity paEntity) {
//        List<XWPFRun> runs = paragraph.getRuns();
//        for (int i = 0; i < runs.size(); ) {
//            XWPFRun run = runs.get(i);
//            if (run == null) {
//                i++;
//                continue;
//            }
//
//            String currentText = run.getText(0);
//            if (StringUtils.isBlank(currentText)) {
//                i++;
//                continue;
//            }
//
//            // 判断是否开始匹配 ${ 或 $ { 跨 run 的情况
//            boolean startsWithDollarBrace = currentText.contains("${");
//            boolean potentialSplit = false;
//
//            if (currentText.contains("$") && i + 1 < runs.size()) {
//                XWPFRun nextRun = runs.get(i + 1);
//                if (nextRun != null) {
//                    String nextText = nextRun.getText(0);
//                    if (StringUtils.isNotBlank(nextText) && nextText.startsWith("{")) {
//                        potentialSplit = true;
//                    }
//                }
//            }
//
//            if (!startsWithDollarBrace && !potentialSplit) {
//                i++;
//                continue;
//            }
//
//            // 开始拼接直到找到右括号 }
//            StringBuilder fullText = new StringBuilder(currentText);
//            while (i + 1 < runs.size() && !fullText.toString().contains("}")) {
//                XWPFRun nextRun = runs.get(i + 1);
//                if (nextRun == null) {
//                    break;
//                }
//                String nextText = nextRun.getText(0);
//                fullText.append(nextText);
//                paragraph.removeRun(i + 1);
//            }
//
//            String finalText = fullText.toString();
//
//            // 获取替换值并设置
//            Object value = changeValuePa(finalText, paEntity);
//            run.setText(value != null ? value.toString() : "", 0);
//
//            i++; // 继续下一个 run
//        }
//    }
//
//
//    /**
//     * 匹配参数
//     *
//     * @param value
//     * @param paEntity
//     * @return
//     */
//    private static Object changeValuePa(String value, PaEntity paEntity) {
//        Object val = "";
//        Map<String, Object> paEntityMap = new HashMap<>();
//        Arrays.stream(BeanUtils.getPropertyDescriptors(DrEntity.class))
//                .filter(pd -> !pd.getName().equals("class"))
//                .forEach(pd -> {
//                    try {
//                        paEntityMap.put(pd.getName(), pd.getReadMethod().invoke(paEntity));
//                    } catch (Exception e) {
//                        // 处理异常
//                    }
//                });
//        for (Map.Entry<String, Object> textSet : paEntityMap.entrySet()) {
//            // 匹配模板与替换值 格式${key}
//            String key = textSet.getKey();
//            if (value.startsWith("${") && value.endsWith("}")) {
//                value = value.substring(2, value.length() - 1);
//            }
//            if (value.equals(key)) {
//                val = textSet.getValue();
//                return val;
//            }
//        }
//        return val;
//    }


}
