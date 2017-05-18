## Excel导入工具类



>  工具基于Apache 的 POI 再度封装，实现从excel到实体的绑定导入。



### 导入原则

主要的基于约定优于配置的这样的习惯来处理解析的。在使用这个工具的时候，我们这样来约定：

1. @ExcelImport 这个注解用来标记需要做导入映射的列，该注解有一个属性 colIndex，对应的就是excel中那一列。而excel的列下标和数组是一样的，从0开始计算。所以，colIndex也是从0开始计算的。没有默认值。
2. 导入的时候，工具类本身不做任何数据类型，数据格式等校验工作，只会把Excel里面的值当成字符串读取出来。具体的校验工作交给实体类的setter方法，如果在setter方法中校验失败了， 那么就抛出异常，工具类会把这一整行当成异常数据进行标记，并重新生成一个错误文件。同时会记录总的错误记录数，追加错误信息后的文件路径，以及解析成功的数据。
3. Excel的格式尽量都是字符串格式的，尽量不出现千奇百怪的日期格式, 如果必须，那么还是建议设置excel为字符串格式，然后在实体类的setter方法做兼容。



### 接口说明

```
public boolean importExcel(File importFile, int startRow)throws Exception {...}
```

导入唯一的入口方法，该方法接收两个参数，一个导入文件，一个是从第几行开始解析文件， startRow的值就是我们在excel看到的是正式值的那一行开始

具体excel解析之后是什么类型，则通过构造函数来说明，同时还会生成excel列和实体字段的映射关系

```
public ExcelImportUtil(Class clazz){
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
```



### 完整代码

```
/**
 * excel解析工具， 用来导入数据
 * 每一个列的校验放到实体类的setter方法来判断，如果验证非法，抛出异常，在解析excel Row的时候捕获异常并在改行的最后列追加一列错误信息描述
 * @author linwb
 * @since 2017-05-18
 */
public class ExcelImportUtil {
    private static final Logger logger = LoggerFactory.getLogger(ExcelImportUtil.class);

    /** 导入的错误文件的位置 后面加上时间戳 ，文件名加上 _error */
    private static final String ERROR_FILE_PATH = "/usr/local/importFiles/";

    /** 导入的实体要映射成何种类型的数据 */
    private Class clazz;
    /** 列->字段名称的映射关系 */
    private Map<String,String> colFieldMap = new HashMap<>();

    /** 成功的记录 */
    private List<Object> successRows = new ArrayList<>();
    /** 错误的行数 */
    private int errorLineCnt = 0;
    /** 错误数据的文件路径 */
    private String errorFilePath = "";

    public int getErrorLineCnt() {
        return errorLineCnt;
    }

    public List<Object> getSuccessRows() {
        return successRows;
    }

    public String getErrorFilePath() {
        return errorFilePath;
    }

    public ExcelImportUtil(Class clazz){
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
            font.setColor(HSSFColor.RED.index);
            cellStyle.setFont(font);


            int rn = sheet.getLastRowNum();
            for(int rowIndex=startRow; rowIndex<rn; rowIndex++){
                Object obj = clazz.newInstance();
                Row row = sheet.getRow(rowIndex);
                try {
                    parseRow2Object(row, obj);
                } catch (Exception e){
                    Cell errCell = row.createCell(row.getLastCellNum());
                    errCell.setCellStyle(cellStyle);
                    // 将错误信息追加到最后一列
                    errCell.setCellValue(e.getCause().getMessage());

                    errorLineCnt++;
                    logger.error("导入的第{}行出错", rowIndex);
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
    private void parseRow2Object(Row row, Object obj) throws Exception {
        Class clazz = obj.getClass();

        boolean canAdd = false;

        int colLen = colFieldMap.keySet().size();
        for(int colIndex=0; colIndex<colLen; colIndex++) {
            Cell cell =  row.getCell(colIndex);
            String cellVal = cell.getStringCellValue();
            if(StringUtils.isBlank(cellVal)){
                continue ;
            }
            String fieldName = colFieldMap.get(StringUtils.EMPTY + colIndex);
            if(StringUtils.isBlank(fieldName)){
                continue ;
            }

            clazz.getMethod("set"+ StringUtil.upperFirstCharName(fieldName), new Class[] {String.class}).invoke(obj, cellVal);

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
//        File f = new File("C:\\Users\\linwb\\Desktop\\新建文件夹\\副本.xls");
//        ExcelImportUtil util = new ExcelImportUtil(SysBlacklist.class);
//        util.importExcel(f, 2);

        File f = new File("C:\\Users\\linwb\\Desktop\\新建文件夹\\副本.xls");
        ExcelImportUtil util = new ExcelImportUtil(SysPyrogeniclist.class);
        util.importExcel(f, 3);

        System.out.println("SUCCESS");
    }
}
```



#### 注解类：

```
/**
 * 导出字段注解
 * 对于 浮点数，日期 等其他特殊字段，可以在要接收导入数据的实体中加上一个set方法来接收字符串，在将字符串转换成其他数据
 * @author linwb
 * @since  2017-05-18
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelImport {

    /**
     * 列的顺序 以 0,1,2,3,4 ... 表示
     */
    String colIndex() ;
}

```



#### Excel需要映射为的实体类 部分内容：

```

public class YourBeanName {
   
    /** 证件号码 */
    private String idcard;
     /** 时限要求 */
    private Float timeLimit;
 
	/** 冗余几个校验字段，用String类型来接收excel的值 */
    @ExcelImport(colIndex = "2")
    private String idcardImport;
    @ExcelImport(colIndex = "9")
    private String timeLimitImport;

    public void setIdcardImport(String idcardImport) throws Exception {
        if(StringUtils.isBlank(idcardImport)){
            return ;
        }
        this.idcardImport = idcardImport;
        if(!IdcardUtil.isValidIdcard(idcardImport, false)) {
            throw new Exception("证件号码非法");
        }
        this.idcard = this.idcardImport;
    }

    public void setTimeLimitImport(String timeLimitImport) throws Exception {
        if(StringUtils.isBlank(timeLimitImport)){
            return ;
        }

        Pattern floatParttern = Pattern.compile("^\\d+(\\.\\d{1})?$", Pattern.CASE_INSENSITIVE);
        boolean b = floatParttern.matcher(timeLimitImport).matches();
        if(!b){
            throw new Exception("时间限制只能为整数或者一位小数点。如: 5 或者 20.5 ");
        }
        this.timeLimit = Float.parseFloat(timeLimitImport);
    }
}
```



