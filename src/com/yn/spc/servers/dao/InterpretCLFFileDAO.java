package com.yn.spc.servers.dao;

import java.io.File;
import java.sql.PreparedStatement;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.regex.Pattern;

import com.yn.spc.common.Keys;
import com.yn.spc.util.FileOptionUtil;
import com.yn.spc.util.ParseUtil;
import com.yn.spc.util.OutRunTimeLog;


public class InterpretCLFFileDAO {
	public int ROU_CLFParam(PreparedStatement ips, File sourceFile, String monitorId,int fDatePOS,String divCode) throws Exception{
		int count=0,colNum=0;
		Pattern pattern = Pattern.compile("\\|");
		String[] patternStr = null;	    
		String encodeType=FileOptionUtil.getTXTEncodeType(sourceFile);
		String[] lines = FileOptionUtil.readLines(sourceFile,encodeType);
		String  StrValue,StrValue1 ="", StrValue2 ="";
		SimpleDateFormat df = new SimpleDateFormat(Keys.DATE_FORMAT2);
		SimpleDateFormat df2 = new SimpleDateFormat(Keys.DATE_FORMAT3);
		SimpleDateFormat df3 = new SimpleDateFormat(Keys.DATE_FORMAT);
		Date dt =null;
		Calendar calFrom = Calendar.getInstance();
		Calendar calTo = Calendar.getInstance();				
		calTo.setTime(df.parse(df.format(new Date())));
		
		try{
			for (String line : lines) {
				if(count<=1){          
					count++;
					continue;
				}		
				
				patternStr = pattern.split(line);
				colNum = patternStr.length ;	
				
				if( colNum>=30 ){
					StrValue1 = ParseUtil.getObjStr(patternStr[16]);
					StrValue2 = ParseUtil.getObjStr(patternStr[24]);					
					
					 if (!"0".equals(StrValue1) && StrValue2.indexOf(".") ==-1 ){
						 
							if(fDatePOS > 0 ){ 
								StrValue = ParseUtil.getObjStr(patternStr[28]);
								
								if("".equals(StrValue)){
								   ips.setDate(1, new java.sql.Date(new Date().getTime()));
								}else{
									dt = new java.sql.Date(df.parse(StrValue).getTime());
									StrValue = df2.format(dt);
									calFrom.setTime(dt);
									calFrom.add(Calendar.DAY_OF_MONTH,52);
									if (calFrom.after(calTo)){
										//showLog("StrValue==: " + StrValue+"\n");	
										ips.setDate(1,new java.sql.Date(df3.parse(StrValue.substring(3,5)+"/"+StrValue.substring(0,2)+"/"+StrValue.substring(6,10)).getTime()) );										
								}else{
										continue;
								     }
								}
							}
							
							ips.setString(2, ParseUtil.getObjStr(patternStr[29]));
							ips.setString(3, ParseUtil.getObjStr(patternStr[24]));
							
							StrValue = ParseUtil.getObjStr(patternStr[2]);
							if(StrValue.indexOf("--") ==-1){
							    ips.setString(4, StrValue);
							}else{								
								ips.setString(4, StrValue.substring(0,StrValue.indexOf("--")));
							}			
							
							ips.setString(5, ParseUtil.getObjStr(patternStr[23]));
							
							StrValue = ParseUtil.getObjStr(patternStr[4]);
							ips.setString(6, StrValue);
							//if(StrValue.indexOf("-") ==-1){
							//    ips.setString(6, StrValue);
							//}else{								
							//	ips.setString(6, StrValue.substring(0,StrValue.indexOf("-")));
							//}
							
							ips.setString(7, ParseUtil.getObjStr(patternStr[17]));
							
							StrValue = ParseUtil.getObjStr(patternStr[16]);						
							if("1".equals(StrValue1)){
								StrValue ="正确";
							}else{
								StrValue ="超出设计值";
							}
							StrValue = " Result为"+StrValue.trim();
							ips.setString(8, StrValue);							
							
							if(fDatePOS == 0){			
								ips.setDate(9, new java.sql.Date(new Date().getTime()));
								ips.setString(10, divCode);
							}else{
								ips.setString(9, divCode);				
							}
							ips.addBatch();
							count=count+1;	
							if(count % 5000 ==0){
								ips.executeBatch();
							}
					 }else{
						 //OutRunTimeLog log = new  OutRunTimeLog();
						 OutRunTimeLog.showLog(monitorId+"数据错误或重复0："+StrValue1+"/"+StrValue2);
					 }
				}
	
			}			
		}catch (Exception e){
			e.printStackTrace();
			//showLog(e.getMessage());		
		}finally{
			return count-2;
		}	
	}
	
	public int CommonCLFParam(PreparedStatement ips, File sourceFile, String monitorId,int fDatePOS,String divCode) throws Exception{
		int count=0,colNum=0;
		Pattern pattern = Pattern.compile("\\|");
		String[] patternStr = null;	    
		String encodeType=FileOptionUtil.getTXTEncodeType(sourceFile);
		String[] lines = FileOptionUtil.readLines(sourceFile,encodeType);
		try{
			for (String line : lines) {
				if(count==0){            //忽略第一行
					count++;
					continue;
				}
				patternStr = pattern.split(line);
				for(colNum=0; colNum<patternStr.length; colNum++){
					if(fDatePOS > 0 && (fDatePOS-colNum)==1){ 
						ips.setDate(colNum+1, new java.sql.Date(new Date().getTime()));
					}else{
						ips.setString(colNum+1, ParseUtil.getObjStr(patternStr[colNum]));
					}					
				}
				if(fDatePOS == 0){			
					ips.setDate(colNum+1, new java.sql.Date(new Date().getTime()));
					ips.setString(colNum+2, divCode);
				}else{
					ips.setString(colNum+1, divCode);				
				}
				ips.addBatch();
				count=count+1;
			}			
		}catch (Exception e){
			e.printStackTrace();
			 OutRunTimeLog.showLog(e.getMessage());		
		}finally{
			return count;
		}	
	}
}
