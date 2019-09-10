package com.mayikt.wordhtml.utils;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

/**
 * @author ChenZhuang
 * @ClassName DateUtil
 * @description TODO
 * @Date 2019/9/9 17:04
 * @Version 1.0
 */
public class DateUtil {
    public static final String YYYY_MM_DD = "yyyy-MM-dd";
    public static final String YYYY_MM_DD_HH_MM_SS = "yyyy-MM-dd HH:mm:ss";
    public static final String YYYYMMDD = "yyyyMMdd";

    public static String date2str(String pattern){
        DateFormat dateFormat = new SimpleDateFormat(pattern);
        return dateFormat.format(new Date());
    }

    public static String date2str(Date date ,String pattern){
        DateFormat dateFormat = new SimpleDateFormat(pattern);
        return dateFormat.format(date);
    }

    //获取前一天日期
    public static Date getYesterDay(){
        Calendar calendar = Calendar.getInstance();
        //这个地方别写成了calendar.set(Calendar.DAY_OF_MONTH,-1);
        calendar.add(Calendar.DAY_OF_MONTH,-1);
        return calendar.getTime();
    }

    //获取前一天日期的字符串
    public static String getYesterDay(String pattern){
        Calendar calendar = Calendar.getInstance();
        //这个地方别写成了calendar.set(Calendar.DAY_OF_MONTH,-1);
        calendar.add(Calendar.DAY_OF_MONTH,-1);
        return date2str(calendar.getTime(),pattern);
    }
}
