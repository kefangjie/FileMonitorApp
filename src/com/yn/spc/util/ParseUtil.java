package com.yn.spc.util;


import java.io.UnsupportedEncodingException;


import java.net.URLDecoder;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Locale;
import java.util.Map;

import net.sf.json.JSONObject;

import org.apache.commons.lang.StringEscapeUtils;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;

import com.yn.spc.common.Keys;



public final class ParseUtil {

	public static String getDateStr() {	
		 return new SimpleDateFormat(Keys.DATE_FORMAT4).format(new Date());		
	}
	public static String getDateStr(Object date) {
		if (date == null) {
			return "";
		} else {
			return new SimpleDateFormat(Keys.DATE_FORMAT).format(date);
		}
	}

	public static String getTimeStampStr(Object date) {
		if (date == null) {
			return "";
		} else {
			return new SimpleDateFormat(Keys.DATE_TIME_FORMAT).format(date);
		}
	}
	
	public static String getEscapeStr(Object val) {
	   String escapeStr = "";
	   if (val != null){
		   escapeStr = StringEscapeUtils.escapeXml(val.toString());
	   }
	   return escapeStr;
   }
	
	public static String getURLStr(Object val) {
	   String urlStr = "";
	   if (val != null){
		   try {
			   urlStr = URLDecoder.decode(StringEscapeUtils.escapeXml(val.toString()),"UTF-8");
		   } catch (UnsupportedEncodingException e) {
			   e.printStackTrace();
		   }
	  }
	  return urlStr;
   }
   
   public static Date getTimeStamp(String dateStr){
	   Date date = null;
	   if (dateStr == null || "".equals(dateStr)) {
		   return null;
	   } else {
		   try {
				date = new SimpleDateFormat(Keys.DATE_TIME_FORMAT).parse(dateStr);
			} catch (ParseException e) {
				e.printStackTrace();
			}
	   }
	   return date;
   }
   
   
   public static Date getDate(String dateStr) {
	   Date date = null;
	   if (dateStr == null || "".equals(dateStr)) {
		   return null;
	   } else {
		   try {
				date = new SimpleDateFormat(Keys.DATE_FORMAT).parse(dateStr);
			} catch (ParseException e) {
				e.printStackTrace();
			}
	   }
	   return date;
   }
   
   public static Map parserToMap(String s){
		Map map=new HashMap();
		if(s!=null && !"".equals(s)){
			JSONObject json=JSONObject.fromObject(s);
			Iterator keys=json.keys();
			while(keys.hasNext()){
				String key=(String) keys.next();
				String value=json.get(key).toString();
				if(value.startsWith("{")&&value.endsWith("}")){
					map.put(key, parserToMap(value));
				}else{
					map.put(key, getURLStr(value));
				}
			}
		}
		return map;
	}
   
   public static String getTrimStr(String val) {
       if (val == null) {
          return ""; 
       }else {
    	   return val.trim();
       }
   }
   
   public static String getObjStr(Object val) {
       if (val == null) {
          return ""; 
       }else {
    	   return val.toString().trim();
       }
   }
   
   public static String getCellValue(Cell cell) {
        String strCell = "";  
        if (cell == null) {  
            return "";  
        }  
        switch (cell.getCellType()) {  
        case Cell.CELL_TYPE_FORMULA: 
            try {  
          	     strCell = String.valueOf(cell.getNumericCellValue());    
          	   //  log.info("strCell=="+strCell);
            } catch (IllegalStateException e) {              
            	  strCell = cell.getRichStringCellValue().getString(); 
            	 // log.info("strCell==="+strCell);
            	  return strCell;  
            }             
            break;  
        case Cell.CELL_TYPE_STRING:  
            strCell = cell.getStringCellValue();  
            break;  
        case Cell.CELL_TYPE_NUMERIC:  
            if (HSSFDateUtil.isCellDateFormatted(cell)) {  
                strCell = cell.getDateCellValue().toString();  
                break;  
            } else {  
                strCell = String.valueOf(cell.getNumericCellValue());  
                break;  
            }  
        case Cell.CELL_TYPE_BOOLEAN:  
            strCell = String.valueOf(cell.getBooleanCellValue());  
            break;  
        case Cell.CELL_TYPE_BLANK:  
            strCell = "";  
            break;  
        default:  
            strCell = "";  
            break;  
        }  
        return strCell;  
    }  

   public static java.sql.Date CellValueToDate(Cell cell) throws ParseException{
	   java.sql.Date date=null;
		if(cell == null ){	
			date = new java.sql.Date(new Date().getTime());
    		  
		}else if(cell.getCellType()==1 || cell.getCellType()==2 ){
			if(cell.getStringCellValue().startsWith("'")){
				cell.setCellValue(cell.getStringCellValue().substring(2));
			}					
			if(cell.getStringCellValue().substring(4, 4).equals("-")){
				date = new java.sql.Date(new SimpleDateFormat("yyyy-MM-dd").parse(cell.getStringCellValue()).getTime());
			}else if(cell.getStringCellValue().substring(4, 4).equals("/")){
				date = new java.sql.Date(new SimpleDateFormat("yyyy/MM/dd").parse(cell.getStringCellValue()).getTime());
			}else if(cell.getStringCellValue().substring(2, 2).equals("-")){
				date = new java.sql.Date(new SimpleDateFormat("MM-dd-yyyy").parse(cell.getStringCellValue()).getTime());
			}else if(cell.getStringCellValue().substring(2, 2).equals("/")){
				date =  new java.sql.Date(new SimpleDateFormat("MM/dd/yyyy").parse(cell.getStringCellValue()).getTime());
			}else{
				date = new java.sql.Date(new Date().getTime());
			}							
		}else if(DateUtil.isCellDateFormatted(cell)){		
			String  DateStr = getCellValue(cell);
			Date dateTime=null;
			SimpleDateFormat formatter = new SimpleDateFormat(Keys.DATE_TIME_FORMAT2);	
			DateFormat df = new SimpleDateFormat(Keys.DATE_TIME_FORMAT3,Locale.US);
			if (DateStr.indexOf("CST") != -1){
				dateTime = df.parse(DateStr);
				DateStr = formatter.format(dateTime);
			}
			//Fri Jun 27 00:00:00 CST 2014
			//2014/06/27 00 00:00
			DateStr = DateStr.replaceAll("-", "/");
			//DateStr = DateStr.replaceFirst(":", " ");
			dateTime=formatter.parse(DateStr);	
			date = new java.sql.Date(dateTime.getTime());
        }  
		
	   return date;
   }
}
