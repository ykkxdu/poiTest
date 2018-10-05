package com.poi.ykk.entity;

import lombok.Data;

/**
 * @Author:Yankaikai
 * @Description:excel表格实体类
 * @Date:Created in 2018/10/4
 */
@Data
public class Student {
    //  学号
    private String num;
    // 姓名
    private String name;
    // 班级
    private String classLevel;
    // 数学
    private String math;
    // 英语
    private String english;
    // 语文
    private String chinese;
}
