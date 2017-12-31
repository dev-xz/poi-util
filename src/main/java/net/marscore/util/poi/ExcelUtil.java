package net.marscore.util.poi;

import java.io.*;
import java.lang.reflect.Field;
import java.util.*;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.DecimalFormat;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static org.apache.poi.ss.usermodel.CellType.STRING;

/**
 * @author Eden
 */
public class ExcelUtil {
    public static List excelFileToList(String fileName, Class tClass) {
        String oldExcelSuffix = ".xls";
        String newExcelSuffix = ".xlsx";
        if (!fileName.toLowerCase().endsWith(oldExcelSuffix)&&!fileName.toLowerCase().endsWith(newExcelSuffix)) {
            throw new RuntimeException("非excel格式文件");
        }
        InputStream excelStream = null;
        try {
            excelStream = new FileInputStream(fileName);
            if (fileName.toLowerCase().endsWith(newExcelSuffix)) {
                return xlsxStreamToList(excelStream, tClass);
            } else if (fileName.toLowerCase().endsWith(oldExcelSuffix)) {
                return xlsStreamToList(excelStream, tClass);
            }
            return null;
        } catch (FileNotFoundException e) {
            throw new RuntimeException("文件不存在",e);
        } finally {
            if (excelStream!=null) {
                try {
                    excelStream.close();
                } catch (IOException e) {
                    // do nothing
                }
            }
        }
    }
    public static List excelFileToList(File file, Class tClass) {
        return excelFileToList(file.getPath(), tClass);
    }

    public static List xlsStreamToList(InputStream excelStream, Class tClass) {
        try {
            Workbook workbook = new HSSFWorkbook(excelStream);
            Sheet sheet = workbook.getSheetAt(0);
            return sheetToList(sheet, tClass);
        } catch (IOException e) {
            throw new RuntimeException("文件损坏，或非excel格式文件",e);
        }
    }

    public static List xlsxStreamToList(InputStream excelStream, Class tClass) {
        try {
            Workbook workbook = new XSSFWorkbook(excelStream);
            Sheet sheet = workbook.getSheetAt(0);
            return sheetToList(sheet, tClass);
        } catch (IOException e) {
            throw new RuntimeException("文件损坏，或非excel格式文件",e);
        }
    }

    public static List sheetToList(Sheet sheet, Class tClass) {
        List<Object> list = new LinkedList<Object>();
        Iterator<Row> rowIterator = sheet.iterator();
        // 获取第一行数据是否与类信息匹配
        if (!rowIterator.hasNext()) {
            throw new RuntimeException("工作表中的内容为空");
        }
        Row row = rowIterator.next();
        Iterator<Cell> cellIterator = row.cellIterator();
        List<String> fields = new LinkedList<String>();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            String value = getCellValue(cell).split("-")[0].trim();
            value = getFieldGetMethodName(value);
            fields.add(value);
        }
        while (rowIterator.hasNext()) {
            row = rowIterator.next();
            cellIterator = row.cellIterator();
            Object t;
            try {
                t = tClass.newInstance();
            } catch (InstantiationException | IllegalAccessException e) {
                throw new RuntimeException("创建实例失败",e);
            }
            for (int i=0; i<fields.size(); i++) {
                String value = null;
                if (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    value = getCellValue(cell);
                }
                Method[] methods = tClass.getMethods();
                for (int j=0; j<methods.length; j++) {
                    if (methods[j].getName().equals("set"+fields.get(i))) {
                        Object realValue = null;
                        if (value!=null && !"".equals(value)) {
                            Class<?> parameterClass = methods[j].getParameterTypes()[0];
                            if (parameterClass.equals(Integer.class)) {
                                value = value.split("\\.")[0];
                            }
                            if (!parameterClass.equals(String.class)) {
                                try {
                                    realValue = parameterClass.getMethod("valueOf", String.class).invoke(null, value);
                                } catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException
                                        | NoSuchMethodException | SecurityException e) {
                                    throw new RuntimeException("转化为与之对应的类型失败",e);
                                }
                            } else {
                                realValue = value;
                            }
                        }
                        try {
                            methods[j].invoke(t, realValue);
                        } catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException e) {
                            throw new RuntimeException("调用方法失败");
                        }
                    }
                }
            }
            list.add(t);
        }
        return list;
    }

    public static OutputStream listToOutputStream(OutputStream outputStream, List list, Map<String, String> userFields) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        listToXSSFSheet(workbook, list, userFields);
        try {
            workbook.write(outputStream);
        } catch (IOException e) {
            throw new RuntimeException("输出流读取错误", e);
        }
        return outputStream;
    }
    public static OutputStream listToOutputStream(OutputStream outputStream, List list) {
        return listToOutputStream(outputStream, list, null);
    }

    public static void listToFile(File file, List list, Map<String, String> userFields) {
        OutputStream outputStream =null;
        try {
            String filePath = file.getPath();
            Integer lastPath = filePath.lastIndexOf("\\");
            if (lastPath==-1) {
                lastPath = filePath.lastIndexOf("/");
            }
            String dirPath = filePath.substring(0, lastPath);
            File dir = new File(dirPath);
            if (!dir.exists()) {
                dir.mkdirs();
            }
            outputStream = new FileOutputStream(file);
            listToOutputStream(outputStream, list, userFields);
            outputStream.flush();
        } catch (IOException e) {
            throw new RuntimeException("文件创建失败", e);
        } finally {
            if (outputStream!=null) {
                try {
                    outputStream.close();
                } catch (IOException e) {
                    // do nothing
                }
            }
        }
    }
    public static void listToFile(File file, List list) {
        listToFile(file, list, null);
    }

    public static void listToFile(String fileName, List list, Map<String, String> userFields) {
        String newExcelSuffix = ".xlsx";
        if (!fileName.toLowerCase().endsWith(newExcelSuffix)){
            fileName = fileName+newExcelSuffix;
        }
        File file = new File(fileName);
        listToFile(file, list, userFields);
    }
    public static void listToFile(String fileName, List list) {
        listToFile(fileName, list, null);
    }

    public static void listToXSSFSheet(XSSFWorkbook workbook, List list, Map<String, String> userFields){
        Class tClass = list.get(0).getClass();
        // 新建工作表
        XSSFSheet sheet = workbook.createSheet(tClass.getSimpleName());
        // 在索引0的位置创建行（最顶端的行）
        XSSFRow row = sheet.createRow(0);
        // 在索引0的位置创建单元格（左上端）
        XSSFCell cell = row.createCell(0);
        cell.setCellType(STRING);
        Field[] fields = tClass.getDeclaredFields();
        List<String> fieldNames = new LinkedList<>();
        for (int i=0; i<fields.length; i++) {
            String alias = null;
            String filedName = fields[i].getName();
            if (userFields!=null && (alias = userFields.get(fields[i].getName()))==null) {
                continue;
            }
            cell = row.createCell(i);
            if (alias!=null && !"".equals(alias)) {
                cell.setCellValue(filedName+"-"+alias);
            } else {
                cell.setCellValue(filedName);
            }
            fieldNames.add(filedName);
        }
        for (int i=0; i<list.size(); i++) {
            row = sheet.createRow(i+1);
            for (int j=0; j<fieldNames.size(); j++) {
                Object value = "";
                try {
                    Method method = tClass.getMethod("get"+getFieldGetMethodName(fieldNames.get(j)));
                    value = method.invoke(list.get(i));
                } catch (NoSuchMethodException e) {
                    throw new RuntimeException("找不到getter方法",e);
                } catch (IllegalAccessException | InvocationTargetException e) {
                    throw new RuntimeException("调用getter方法失败",e);
                }
                cell = row.createCell(j);
                cell.setCellValue(value.toString());
            }
        }
        for (int i=0; i<fields.length; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private static String getCellValue(Cell cell) {
        Object result = "";
        if (cell != null) {
            switch (cell.getCellTypeEnum()) {
                case STRING:
                    result = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    result = cell.getNumericCellValue();
                    result = new DecimalFormat("#").format(result);
                    break;
                case BOOLEAN:
                    result = cell.getBooleanCellValue();
                    break;
                case FORMULA:
                    result = cell.getCellFormula();
                    break;
                case ERROR:
                    result = cell.getErrorCellValue();
                    break;
                case BLANK:
                    break;
                default:
                    break;
            }
        }
        return result.toString();
    }
    private static String getFieldGetMethodName(String str){
        byte[] items = str.getBytes();
        int ASCIIOfA = (int)'A';
        int ASCIIOfB = (int)'Z';
        if ((items[0] >= ASCIIOfA) && (items[0] <= ASCIIOfB)) {
            return str;
        }
        items[0] = (byte) ((char) items[0] - 'a' + 'A');
        return new String(items);
    }
}