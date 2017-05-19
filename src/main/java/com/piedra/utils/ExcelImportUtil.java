package com.piedra.utils;

import com.piedra.annotation.ExcelImport;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.File;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.util.*;

/**
 * excel解析工具， 用来导入数据
 * 每一个列的校验放到实体类的setter方法来判断，如果验证非法，抛出异常，在解析excel Row的时候捕获异常并在改行的最后列追加一列错误信息描述
 * @author linwb
 * @since 2017-05-18
 */
public class ExcelImportUtil<T> {
    private static final Logger logger = LoggerFactory.getLogger(ExcelImportUtil.class);

    /** 导入的错误文件的位置 后面加上时间戳 ，文件名加上 _error /usr/local/importFiles */
    private static final String ERROR_FILE_PATH = "C:\\Users\\Administrator\\Desktop\\importFiles\\";

    /** 导入的实体要映射成何种类型的数据 */
    private Class<T> clazz;
    /** 列->字段名称的映射关系 */
    private Map<String,String> colFieldMap = new HashMap<>();


    /** 成功的记录 */
    private List<T> successRows = new ArrayList<>();
    /** 错误的行数 */
    private int errorLineCnt = 0;
    /** 错误数据的文件路径 */
    private String errorFilePath = "";

    public int getErrorLineCnt() {
        return errorLineCnt;
    }

    public List<T> getSuccessRows() {
        return successRows;
    }

    public String getErrorFilePath() {
        return errorFilePath;
    }

    public ExcelImportUtil(Class<T> clazz){
        this.clazz = clazz;

        // 通过解析Clazz的annotation来映射 excel列和字段的映射关系
        Field[] fields = clazz.getDeclaredFields();
        for(Field field : fields){
            ExcelImport anno = field.getAnnotation(ExcelImport.class);
            if(anno==null){
                continue ;
            }
            String colIndex = anno.colIndex();
            if(colFieldMap.containsKey(colIndex)){
                throw new RuntimeException("导入字段列顺序配置重复了 [ " + colIndex + " ]");
            }
            colFieldMap.put(colIndex, field.getName());
        }
    }

    /**
     * 预导入验证
     * @throws Exception    抛出异常
     */
    private void preValidate() throws Exception {
        if(clazz==null){
            throw new Exception("请通过有参的构造函数来构造工具类: ExcelImportUtil(Class clazz) ");
        }
    }


    /***
     * 具体的格式转换会交给 clazz 类型的setter方法处理， 可以另外起一个字段专门处理这种类型的转换工作，并抛出异常
     * @param importFile    导入文件
     * @param startRow      开始解析的行下标
     * @return  返回解析是否成功
     * @throws Exception
     */
    public boolean importExcel(File importFile, int startRow)throws Exception {
        preValidate();

        if(startRow<0){
            startRow = 1;
        }
        startRow--; // 理论上第1行，对Sheet来说是第0行

        errorFilePath = getErrorFilePath(importFile.getName());
        FileOutputStream out = null;
        try {
            Workbook wb = WorkbookFactory.create(importFile);
            Sheet sheet = wb.getSheetAt(0);

            CellStyle cellStyle = wb.createCellStyle();
            Font font = wb.createFont();
            font.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
            cellStyle.setFont(font);


            int rn = sheet.getLastRowNum();
            for(int rowIndex=startRow; rowIndex<rn; rowIndex++){
                T obj = clazz.newInstance();
                Row row = sheet.getRow(rowIndex);
                try {
                    parseRow2Object(row, obj);
                } catch (Exception e){
                    Cell errCell = row.createCell(row.getLastCellNum());
                    errCell.setCellStyle(cellStyle);
                    // 将错误信息追加到最后一列
                    String errorMsg = e.getMessage();
                    if(StringUtils.isBlank(errorMsg) && e.getCause()!=null){
                        errorMsg = e.getCause().getMessage();
                    }
                    if(StringUtils.isBlank(errorMsg)){
                        errorMsg = "导入出错";
                    }
                    errCell.setCellValue(errorMsg);

                    errorLineCnt++;
                    logger.error("导入的第{}行出错", rowIndex, e);
                }
            }

            // 再把结果写到另一个错误文件中
            out = new FileOutputStream(new File(errorFilePath));
            wb.write(out);

        } catch (Exception e){
            logger.error("导入Excel失败", e);
            return false;
        } finally {
            IOUtils.closeQuietly(out);
        }
        return true;
    }

    /**
     * 将excel一行转换成clazz对象
     * @param row   excel行数据
     * @throws Exception    抛出异常
     */
    @SuppressWarnings("unchecked")
    private void parseRow2Object(Row row, T obj) throws Exception {
        Class clazz = obj.getClass();
        boolean canAdd = false;
        int colLen = colFieldMap.keySet().size();
        for(int colIndex=0; colIndex<colLen; colIndex++) {
            Cell cell =  row.getCell(colIndex);
            cell.setCellType(CellType.STRING);
            String cellVal = cell.getStringCellValue();
            if(StringUtils.isBlank(cellVal)){
                continue ;
            }
            String fieldName = colFieldMap.get(StringUtils.EMPTY + colIndex);
            if(StringUtils.isBlank(fieldName)){
                continue ;
            }
            clazz.getMethod("set"+Character.toUpperCase(fieldName.charAt(0))+fieldName.substring(1), new Class[]{String.class}).invoke(obj, cellVal);
            canAdd = true;
        }
        if(canAdd) {
            successRows.add(obj);
        }
    }


    /**
     * 获取错误信息文件的路径
     * @param fileName  源文件名
     * @return  返回错误信息
     */
    private String getErrorFilePath(String fileName){
        Calendar c = Calendar.getInstance();
        c.setTime(new Date());

        String dir = ERROR_FILE_PATH+"error/" + c.get(Calendar.YEAR)+c.get(Calendar.MONTH)+c.get(Calendar.DAY_OF_MONTH)+c.getTimeInMillis();
        File fileDir = new File(dir);
        if(!fileDir.exists()){
            fileDir.mkdirs();
        }
        int dotIndex = fileName.lastIndexOf(".");
        return dir + File.separator + fileName.substring(0, dotIndex) + "_error" + fileName.substring(dotIndex);
    }


    public static void main(String[] args) throws  Exception{
//        File f = new File("C:\\Users\\Administrator\\Desktop\\名单.xls");
//        ExcelImportUtil<User> util = new ExcelImportUtil<>(User.class);
//        util.importExcel(f, 2);
//
//        System.out.println("SUCCESS");



    }

}
