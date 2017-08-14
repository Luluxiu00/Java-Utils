package com.deppon.dpap.module.base.shared.utils;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import org.apache.commons.lang3.StringUtils;

/**
 *  日期操作工具类
 * Created by WZ on 2017-5-5.
 * empCode: 326944
 */
public class DateUtils {

	/**
	 * 日期添加下划线
	 * @param date
	 * @return
	 */
	public static String dateAdd_(String date){
		String year=date.substring(0,4);
		String month=date.substring(4,6);
		String day=date.substring(6,8);
		return year+"-"+month+"-"+day;
	}
	/**
	 * 判断参数 时间 是否是 当天日期
	 * @param date	
	 * @return	返回参数 时间戳是否是当前日期
	 */
	public static boolean isCurrentDay( String dateStr, String pattern){
		if( dateStr == null || dateStr.isEmpty()){
			return false;
		}
		return isCurrentDay( string2Date(dateStr, pattern));
	}
	/**
	 * 判断参数 时间 是否是 当天日期
	 * @param date	
	 * @return	返回参数 时间戳是否是当前日期
	 */
	public static boolean isCurrentDay( Date date){
		if( date == null)
			return false;
		return getCurrentDateString("yyyyMMdd").equals( date2String( date, "yyyyMMdd"));
	}
	
	/**
	 * 获得时间差
	 */
	public static int getTimeDifference(){
		long startTime=DateUtils.getFixedTime(8).getTime()/1000;
		long endTime=System.currentTimeMillis()/1000;
		int minute=(int) ((endTime-startTime)/60);
		return minute;
	}

	/**
	 * 根据时间获得当前时间
	 * @return
	 */
	public static String getStrByDate(Date date){
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		String currTime=sdf.format(date);
		return currTime;
	}
	/**
	 * 获得前日期
	 * @return
	 */
	public static String getBeforeDate(int num){
		Calendar cal = Calendar.getInstance();
		cal.add(Calendar.DATE, num);
		String month=(cal.get(Calendar.MONTH)+ 1)+"";
		month=month.length()==1 ? "0"+month : month;
		String day=(cal.get(Calendar.DAY_OF_MONTH))+"";
		day=day.length()==1 ? "0"+day : day;
		String locusDate = cal.get(Calendar.YEAR) + "" + month+ "" + day;
		return locusDate;
	}
	/**
	 * 获得当前日期
	 * @return
	 */
	public static String getCurrDate(){
		Calendar cal = Calendar.getInstance();
		String month=(cal.get(Calendar.MONTH)+ 1)+"";
		month=month.length()==1 ? "0"+month : month;
		String day=(cal.get(Calendar.DAY_OF_MONTH))+"";
		day=day.length()==1 ? "0"+day : day;
		String locusDate = cal.get(Calendar.YEAR) + "" + month+ "" + day;
		return locusDate;
	}

	/**
	 * 获得每天固定时间段
	 * @return
	 */
	public static Date getFixedTime(int hour) {
		Calendar cal = Calendar.getInstance();
        // 每天定点执行
        cal.set(Calendar.HOUR_OF_DAY, hour);
        cal.set(Calendar.MINUTE, 0);
        cal.set(Calendar.SECOND, 0);
		return cal.getTime();
	}
	/**
	 * @param pattern 日期格式，若为空，则默认“yyyy-MM-dd hh:mm:ss SSS”
	 * @return 根据pattern格式化当前日期，并返回
	 */
	public static String getCurrentDateString(String pattern){
		return date2String(new Date(), pattern);
	}
	/**
	 * @return 返回当前日期，时间为00:00:00
	 */
	public static Date getCurrentDate(){
		return string2Date(new SimpleDateFormat("yyyy-MM-dd").format(new Date()),"yyyy-MM-dd");
	}
	/**
	 * @param dateStr 日期字符串
	 * @param pattern	日期字符串格式
	 * @return 根据pattern解析dateStr，返回解析结果
	 */
	public static Date string2Date(String dateStr, String pattern){
		if(StringUtils.isNotEmpty(dateStr) && StringUtils.isNotEmpty(pattern)){
			try {
				return new SimpleDateFormat(pattern).parse(dateStr);
			} catch (ParseException e) {
				return null;
			}
		}
		return null;
	}
	
	/**
	 * @param date	日期
	 * @param pattern	日期格式，若pattern为空，则默认“yyyy-MM-dd hh:mm:ss SSS”
	 * @return	将日期按pattern转化为字符串
	 */
	public static String date2String(Date date, String pattern){
		if(date != null){
			if(StringUtils.isEmpty(pattern)){
				pattern = "yyyy-MM-dd hh:mm:ss SSS";
			}
			return new SimpleDateFormat(pattern).format( date);
		}
		return null;
	}
	
	/**
	 * @param date	参照日期
	 * @param month 改变数量，可以为负
	 * @return	返回参照日期 加 n个月的日期
	 */
	public static Date addMonths( Date date, int month){
		Calendar c = Calendar.getInstance();
		c.setTime(date);
		c.add(Calendar.MONTH, month);
		return c.getTime();
	}
	/**
	 * @param date 参照日期
	 * @param day 改变数量，可以为负
	 * @return	返回参照日期 加 n天 的日期
	 */
	public static Date addDays( Date date, int day){
		Calendar c = Calendar.getInstance();
		c.setTime(date);
		c.add(Calendar.DAY_OF_MONTH, day);
		return c.getTime();
	}
	
	/**
	 * @param date 参照日期
	 * @param month 改变数量，可以为负
	 * @return	返回参照日期 加 n月 的第一天
	 */
	public static Date firstDay( Date date, int month){
		Calendar c = Calendar.getInstance();    
		c.add(Calendar.MONTH, month);
		c.set(Calendar.DAY_OF_MONTH,1);//设置为1号,当前日期既为本月第一天 
		return c.getTime();
	}

	/**
	 * 获取从当前到指定时间点的 秒 数
	 * @param date
	 * @return 返回 从现在到指定时间点 的 秒 数。
	 * 			当指定时间点为null时返回null；
	 * 			当指定时间点在当前时间点之前时，返回秒数为负数
	 * 			当指定时间点在当前时间点之后时，返回秒数为正数
	 */
	public static Integer getSecsDifference(Date date){
		if(date == null) 
			return null;
		return (int)((date.getTime() - new Date().getTime())/1000);
	}
	
	/**
	 * 根据系统当前时间 获取下一个半点 的 时间点
	 * @return 返回下一个半点
	 * 		若当前时间为 2:10，则返回 2:30
	 * 		若当前时间为 2:40，则返回 3:00
	 */
	public static Date getNextHalfHour() {
		Calendar c = Calendar.getInstance();
		c.setTime(new Date());
		int minute = c.get(Calendar.MINUTE);
		if(minute > 30){
			c.add(Calendar.HOUR_OF_DAY, 1);
			c.set(Calendar.MINUTE, 0);
			c.set(Calendar.SECOND, 0);
		}else if(minute < 30){
			c.set(Calendar.MINUTE, 30);
			c.set(Calendar.SECOND, 0);
		}
		return c.getTime();
	}
	
	/**
	 * 获取指定日期的前第几天
	 * @param specifiedDay 日期
	 * @param num 前几天
	 * @return
	 */
	public static String getBeforeDay2(String specifiedDay,int num){  
        Calendar c = Calendar.getInstance();  
        Date date = null;  
        try{  
            date = new SimpleDateFormat("yyyyMMdd").parse(specifiedDay);  
        }catch (Exception e) {  
            e.printStackTrace();  
        }  
        c.setTime(date);  
        int day = c.get(Calendar.DATE);  
        c.set(Calendar.DATE, day-num);  
        String dayBefore = new SimpleDateFormat("yyyyMMdd").format(c.getTime());  
        return dayBefore;  
    }  
}
