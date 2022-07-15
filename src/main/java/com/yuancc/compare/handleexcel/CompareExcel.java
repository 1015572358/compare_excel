package com.yuancc.compare.handleexcel;

import com.yuancc.compare.controller.MyController;
import com.yuancc.compare.domain.Wgtxz;
import com.yuancc.compare.utils.ExcelUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.nio.charset.Charset;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/**
 * @author ycc
 */
@RestController
public class CompareExcel {
    private static Logger log = Logger.getLogger(CompareExcel.class);
    //目前路径.zip
    public static final String targetSrc = "D://compare";
    //资源路径.xlsx
    public static final String sourceSrc = "D://compare";

    static Map<String, Object> TargetMap = new HashMap<>();
    static LinkedList<Wgtxz> TargetList = new LinkedList<>();

    @GetMapping("/compare")
    public static String getCompareExcel(HttpServletResponse response) {
        File allFiles = new File(targetSrc);
        File[] tempList = allFiles.listFiles();
        log.info("开始校验文件夹内的文件是否存在");
        if (tempList == null) {
            log.info( "请在文件夹下放入要比对的文件");
            return "请在文件夹下放入要比对的文件";
        } else {
            boolean zipExit = false;
            boolean xlsxExit = false;
            for (int i = 0; i < tempList.length; i++) {
                if (tempList[i].isFile()) {
                    if (tempList[i].getName().endsWith(".rar") || tempList[i].getName().endsWith(".zip")) {
                        zipExit = true;
                    } else if (tempList[i].getName().endsWith(".xlsx") || tempList[i].getName().endsWith(".xls")) {
                        xlsxExit = true;
                    }
                }
            }
            if (zipExit && xlsxExit) {
                log.info("校验完毕，文件存在");
                //清空map
                TargetMap = new HashMap<>();
                TargetList = new LinkedList();
                for (int i = 0; i < tempList.length; i++) {
                    if (tempList[i].isFile()) {
                        if (tempList[i].getName().endsWith(".rar") || tempList[i].getName().endsWith(".zip")) {
                            zipDecompress(tempList[i]);
                            log.info("zip包处理完毕");
                            log.info("获取到的数量:" + TargetMap.keySet().size());
                        }
                    }
                }
                for (int i = 0; i < tempList.length; i++) {
                    if (tempList[i].getName().endsWith(".xlsx") || tempList[i].getName().endsWith(".xls")) {
                        log.info("xlsx开始处理");
                        log.info("开始读取xls:"+tempList[i].getName());
                        LinkedList<Wgtxz> wgtxzs = ExcelUtils.readExcel(Wgtxz.class, tempList[i], 1);
                        log.info("xlsx数据读取完毕，记录在list临时集合中数量为：" + wgtxzs.size());
                        log.info("放入公共集合。。。");
                        TargetList.addAll(wgtxzs);
                    }
                }

                log.info("开始匹配人员");
                Iterator<Wgtxz> itr = TargetList.iterator();
                StringBuilder sb = new StringBuilder(40);
                while (itr.hasNext()) {
                    Wgtxz wgtxz = itr.next();
                    if ("".equals(wgtxz.getXm()) || "".equals(wgtxz.getSfzh())) {
                        continue;
                    }
                    if (wgtxz.getSfzh().length() != 18 ) {
                        continue;
                    }
                    sb.append(wgtxz.getXm());
                    sb.append(wgtxz.getSfzh().substring(0, 10));
                    sb.append(wgtxz.getSfzh().substring(14, 18));
//                            sb.append(wgtxz.getPhone().substring(0, 3));
//                            sb.append(wgtxz.getPhone().substring(5, 11));
                    if (TargetMap.containsKey(sb.toString())) {
                        itr.remove();
                    }
//                                else {
//                                    log.info("未匹配了相同人员姓名：" + wgtxz.getXm());
//                                }
                    //重置sb
                    sb.delete(0,sb.length());
                }
                log.info("比对结束：未匹配人员数量：" + TargetList.size());
                log.info("开始导出人员信息。。。");
                ExcelUtils.writeExcel(response, TargetList, Wgtxz.class);

            }
            log.info("成功结束");
            return "成功结束";
        }
    }

    public static void zipDecompress(File file) {
        log.info("开始解压zip包");
        try {
            ZipFile zipFile = new ZipFile(file, Charset.forName("GBK"));
            Enumeration<? extends ZipEntry> enumeration = zipFile.entries();
            ZipEntry zipEntry;
            while (enumeration.hasMoreElements()) {
                zipEntry = (ZipEntry) enumeration.nextElement();
                log.info("解压文件名为:" + zipEntry.getName());
                if (zipEntry.isDirectory()) {
                    continue;
                }
                InputStream is = zipFile.getInputStream(zipEntry);
                PutKeyToMap(is, zipEntry.getName());
                log.info("map装载完毕！");
//                FileOutputStream fos = new FileOutputStream(xlsxfile);
//                int len;
//                byte[] bytes = new byte[1024];
//                while ((len=is.read())!=-1){
//                    fos.write(len);
//                }
//                fos.close();
                is.close();
            }
            zipFile.close();

        } catch (IOException e) {
            e.printStackTrace();
        }


    }


    public static void PutKeyToMap(InputStream is, String fileName) {

        log.info("开始读取文件：" + fileName);
        Workbook workbook = null;
        InputStream inputStream = null;
        try {
            inputStream = is;
            if (fileName.endsWith("xlsx")) {
                workbook = new XSSFWorkbook(inputStream);
            }
            if (fileName.endsWith("xls")) {
                workbook = new HSSFWorkbook(inputStream);
            }
            if (workbook != null) {
                //默认一个sheet页面
                for (int m = 0; m < workbook.getNumberOfSheets(); m++) {
                    Sheet sheet = workbook.getSheetAt(m);
                    //获取所有合并的单元格
                    StringBuilder sb = new StringBuilder();
                    log.info("开始写入临时map集合");
                    for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
                        Row row = sheet.getRow(i);
                        //忽略空白行
                        if (row == null) {
                            continue;
                        }
                        try {
                            //判断是否为空白行
                            boolean allBlank = true;
                            //姓名
                            String xm = getCellValue(row.getCell(7));
                            //身份证号
                            String sfzh = getCellValue(row.getCell(8));
                            //手机号
//                        String phone = getCellValue(row.getCell(10));
                            if (null == xm || "".equals(xm.replace(" ", "")) || null == sfzh || "".equals(sfzh.replace(" ", ""))) {
                                continue;
                            }
                            //默认18位
                            if (sfzh.replace(" ", "").length() != 18) {
                                continue;
                            }
                            //默认11位
//                        if (phone.replace(" ", "").length() != 11) {
//                            continue;
//                        }

                            sb.append(xm.replace(" ", ""));
                            sb.append(sfzh.replace(" ", "").substring(0, 10));
                            sb.append(sfzh.replace(" ", "").substring(14, 18));
//                        sb.append(phone.replace(" ", "").substring(0, 3));
//                        sb.append(phone.replace(" ", "").substring(5, 11));
                            TargetMap.put(sb.toString(), xm);
                            //清空数据
                            sb.delete(0, sb.length());
                        } catch (Exception e) {
                            e.printStackTrace();
                            log.error(String.format("parse row:%s exception!", i), e);
                        }
                    }
                }
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    public static List<CellRangeAddress> getCombineCell(Sheet sheet) {
        List<CellRangeAddress> list = new ArrayList<>();
        //获得一个 sheet 中合并单元格的数量
        int sheetmergerCount = sheet.getNumMergedRegions();
        //遍历所有的合并单元格
        for (int i = 0; i < sheetmergerCount; i++) {
            //获得合并单元格保存进list中
            CellRangeAddress ca = sheet.getMergedRegion(i);
            list.add(ca);
        }
        return list;
    }


    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            if (HSSFDateUtil.isCellDateFormatted(cell)) {
                return HSSFDateUtil.getJavaDate(cell.getNumericCellValue()).toString();
            } else {
                return new BigDecimal(cell.getNumericCellValue()).toString();
            }
        } else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
            return StringUtils.trimToEmpty(cell.getStringCellValue());
        } else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
            return StringUtils.trimToEmpty(cell.getCellFormula());
        } else if (cell.getCellType() == Cell.CELL_TYPE_BLANK) {
            return "";
        } else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        } else if (cell.getCellType() == Cell.CELL_TYPE_ERROR) {
            return "ERROR";
        } else {
            return cell.toString().trim();
        }

    }

}
