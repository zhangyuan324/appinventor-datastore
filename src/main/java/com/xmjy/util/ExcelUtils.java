package com.xmjy.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.net.URLEncoder;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.regex.Pattern;

public class ExcelUtils {

    private static NumberFormat numberFormat = NumberFormat.getInstance();

    static {
        numberFormat.setGroupingUsed(false);
    }

    private static String getRealStringValueOfDouble(Double d) {
        String doubleStr = d.toString();
        boolean b = doubleStr.contains("E");
        int indexOfPoint = doubleStr.indexOf('.');
        if (b) {
            int indexOfE = doubleStr.indexOf('E');
            BigInteger xs = new BigInteger(doubleStr.substring(indexOfPoint
                    + BigInteger.ONE.intValue(), indexOfE));
            int pow = Integer.valueOf(doubleStr.substring(indexOfE
                    + BigInteger.ONE.intValue()));
            int xsLen = xs.toByteArray().length;
            int scale = xsLen - pow > 0 ? xsLen - pow : 0;
            doubleStr = String.format("%." + scale + "f", d);
        } else {
            Pattern p = Pattern.compile(".0$");
            java.util.regex.Matcher m = p.matcher(doubleStr);
            if (m.find()) {
                doubleStr = doubleStr.replace(".0", "");
            }
        }
        return doubleStr;
    }

    /**
     * 设置下载请求头和文件名
     *
     * @param request  request
     * @param response response
     * @param fileName 文件名称
     */
    private static void setResponseHeader(HttpServletRequest request, HttpServletResponse response, String fileName) {
        try {
            try {
                if (request.getHeader("USER-AGENT").toLowerCase().contains("firefox")) {
                    response.setCharacterEncoding("utf-8");
                    response.setHeader("content-disposition",
                            "attachment;filename=" + new String(fileName.getBytes(), "ISO8859-1") + ".xls");
                } else {
                    response.setCharacterEncoding("utf-8");
                    response.setHeader("content-disposition",
                            "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8") + ".xlsx");
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
            response.setContentType("application/msexcel;charset=UTF-8");
            response.setHeader("Pragma", "No-cache");
            response.setHeader("Cache-Control", "no-cache");
            response.setDateHeader("Expires", 0);
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    /**
     * 导出excel
     *
     * @param req
     * @param res
     * @param headers    表头
     * @param properties 表头对应obj属性名(注：与表头字段顺序一致)
     * @param fileName   文件名
     * @param sheetName  表名
     * @param obj        不定参数，现在只接受一个excel要显示的Model的一个List
     * @return
     */
    public static void download(HttpServletRequest req, HttpServletResponse res, String fileName, String sheetName,
                                String[] headers, String properties[], Object... obj) {
        // 设置下载请求头和文件名
        setResponseHeader(req, res, fileName);
        // 声明一个工作薄
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 生成一个表格
        XSSFSheet sheet = workbook.createSheet(sheetName);
        for (int i = 0; i < headers.length; i++) {
            sheet.setColumnWidth(i, 6000);
        }

        XSSFRow row = sheet.createRow(0);
        // 创建表头
        for (short i = 0; i < headers.length; i++) {
            XSSFCell cell = row.createCell(i);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);

        }
        if (properties != null) {
            List<Object> list = (List<Object>) obj[0]; // 获取目标参数（即需要导出的数据）
            for (int j = 0; j < list.size(); j++) {
                XSSFRow r = sheet.createRow(j + 1);
                row.setHeightInPoints(20);
                Object o = list.get(j);
                for (int i = 0; i < properties.length; i++) {
                    String property = properties[i];
                    String getMethodName = "get" // 拼接该字段对应的get方法的方法名
                            + property.substring(0, 1).toUpperCase() + property.substring(1);
                    Class cls = o.getClass();
                    try {
                        Method getMethod = cls.getMethod(getMethodName, new Class[]{});
                        try {
                            Object target = getMethod.invoke(o, new Object[]{}); // 利用Java反射机制调用get方法获取字段对应的值
                            XSSFCell c = r.createCell(i);
                            if (target instanceof Integer) {
                                c.setCellValue((Integer) target);
                            } else if (target instanceof String) {
                                c.setCellValue(target.toString());
                            } else if (target instanceof LocalDateTime) {
                                DateTimeFormatter dateFormat = DateTimeFormatter.ofPattern("yyyy/MM/dd HH:mm:ss");
                                String readDateValue = ((LocalDateTime) target).format(dateFormat);
                                c.setCellValue(readDateValue);
                            } else if (target instanceof Double) {
                                c.setCellValue(decimal((Double) target)); // 如果target为Double类型的数据，则默认保留两位小数
                            } else if (target instanceof Long) {
                                c.setCellValue((Long) target);
                            } else if (target instanceof BigDecimal) {
                                c.setCellValue(decimal(((BigDecimal) target).doubleValue()));
                            }
                        } catch (IllegalAccessException e) {
                            System.out.println(e.getMessage());
                        } catch (IllegalArgumentException e) {
                            System.out.println(e.getMessage());
                        } catch (InvocationTargetException e) {
                            System.out.println(e.getMessage());
                        }
                    } catch (NoSuchMethodException e) {
                        System.out.println(e.getMessage());
                    } catch (SecurityException e) {
                        System.out.println(e.getMessage());
                    }
                }
            }
        }
        OutputStream ouputStream;
        try {
            ouputStream = res.getOutputStream();
            workbook.write(ouputStream);
            ouputStream.flush();
            ouputStream.close();
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
    }

    /**
     * 创建一个工作簿
     *
     * @param is
     * @param excelFileName
     * @return
     * @throws IOException
     */
    public static Workbook createWorkbook(InputStream is, String excelFileName) throws IOException {
        if (excelFileName.endsWith(".xls")) {
            return new HSSFWorkbook(is);
        } else if (excelFileName.endsWith(".xlsx")) {
            return new XSSFWorkbook(is);
        }
        return null;
    }

    /**
     * @Description: 根据sheet索引号获取对应的sheet
     */
    public static Sheet getSheet(Workbook workbook, int sheetIndex) {
        return workbook.getSheetAt(0);
    }

    /**
     * 将Date型转为Date型
     *
     * @param date
     * @return format ||yyyy-MM-dd HH:mm:ss
     */
    public static Date convertDateToDate(Date date) {
        String format = "yyyy-MM-dd HH:mm:ss";
        if (date != null) {
            try {
                SimpleDateFormat sdf = new SimpleDateFormat(format);
                return sdf.parse(sdf.format(date));
            } catch (ParseException e) {
                e.printStackTrace();
            }
        }
        return null;
    }

    /**
     * 保留两位小数
     *
     * @param v
     * @return
     */
    public static double decimal(double v) {
        BigDecimal b = new BigDecimal(Double.toString(v));
        BigDecimal one = new BigDecimal("1");
        return b.divide(one, 2, 1).doubleValue();
    }

    //构建对象
    public static <T> T create(Class<T> clazz) {
        Object t = null;
        try {
            t = clazz.newInstance();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return (T) t;
    }
}
