## Excel导入工具类



>  工具基于Apache 的 POI 再度封装，实现从excel到实体的绑定导入。



### 导入原则

主要的基于约定优于配置的这样的习惯来处理解析的。在使用这个工具的时候，我们这样来约定：

1. @ExcelImport 这个注解用来标记需要做导入映射的列，该注解有一个属性 colIndex，对应的就是excel中那一列。而excel的列下标和数组是一样的，从0开始计算。所以，colIndex也是从0开始计算的。没有默认值
2. 导入的时候，工具类本身不做任何数据类型，数据格式等校验工作，只会把Excel里面的值当成字符串读取出来。具体的校验工作交给实体类的setter方法，如果在setter方法中校验失败了， 那么就抛出异常，工具类会把这一整行当成异常数据进行标记，并重新生成一个错误文件。同时会记录总的错误记录数，追加错误信息后的文件路径，以及解析成功的数据。
3. Excel的格式在导入解析的时候都会被当成字符串的形式，如果希望解析成其他具体类型的数据，如日期格式，浮点数，整数 等 都可以在实体类的setter中进行校验



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



### 注解类：

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



### Excel需要映射为的实体类 部分内容：

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



