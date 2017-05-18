package com.piedra.bean;

import com.piedra.annotation.ExcelImport;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateUtils;

import java.util.Date;

/**
 * @author webinglin
 * @since 2017-05-18
 */
public class User {
    /** 用户名，任意 */
    @ExcelImport(colIndex = "0")
    private String username;

    /** 整数 */
    private int age;

    /** 年-月-日 格式 */
    private Date birthday;

    /* *************************    需要做校验的字段，只提供setter方法，统一用字符串格式接收再去转换  ******************************** */

    @ExcelImport(colIndex = "1")
    private String ageImport;
    @ExcelImport(colIndex = "2")
    private String birthdayImport;

    public void setAgeImport(String ageImport) throws Exception {
        if(StringUtils.isBlank(ageImport)){
            return ;
        }

        // 必须是整数
        if(!ageImport.matches("[1-9]\\d*")) {
            throw new Exception("年龄必须是整数");
        }

        this.age = Integer.parseInt(ageImport);
    }

    public void setBirthdayImport(String birthdayImport) throws Exception {
        if(StringUtils.isBlank(birthdayImport)){
            return ;
        }

        this.birthday = DateUtils.parseDate(birthdayImport, "yyyy-MM-dd");
    }

    public String getUsername() {
        return username;
    }

    public void setUsername(String username) {
        this.username = username;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public Date getBirthday() {
        return birthday;
    }

    public void setBirthday(Date birthday) {
        this.birthday = birthday;
    }
}
