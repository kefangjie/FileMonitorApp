package com.yn.spc.servers.dao;


import java.io.File;
import java.math.BigDecimal;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Timestamp;
import java.text.DateFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Locale;

import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;


import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.hibernate.Query;
import org.hibernate.SQLQuery;
import org.hibernate.Session;

import com.yn.spc.util.DBOptionUtil;
import com.yn.spc.util.FileOptionUtil;
import com.yn.spc.util.OutRunTimeLog;
import com.yn.spc.util.ParseUtil;
import com.yn.spc.common.Keys;

public class InterpretXLSFileDAO {
        
		private void setIPSValue1000(PreparedStatement ips,Cell cell, int postion) throws SQLException{
			String realValue="";
			if(cell!=null){
				if(cell.getCellType()==1){    
					realValue=cell.getStringCellValue();
				}else if(cell.getCellType()==0){  
					realValue=cell.getNumericCellValue()+"";  //realValue=cell.getNumericCellValue()*1000+"";
					if(realValue.endsWith(".0")){
						realValue=realValue.substring(0,realValue.length()-2);
					}else if(realValue.contains(".")){
						realValue=getRoundNum(realValue);
					}
				}else if(cell.getCellType()==2){
					realValue=cell.getNumericCellValue()+"";
					if(realValue.endsWith(".0")){
						realValue=realValue.substring(0,realValue.length()-2);
					}else if(realValue.contains(".")){
						realValue=getRoundNum(realValue);
					}
				}else{
					realValue="0";
				}
			}else{
				realValue="0";
			}
			realValue = realValue.replaceAll(",", "");
			ips.setString(postion, realValue);
		}
		
		private void setIPSValueForInner(PreparedStatement ips,Cell cell, int postion) throws SQLException{
			String realValue="";
			if(cell!=null){
				if(cell.getCellType()==1){    
					realValue=cell.getStringCellValue();
				}else if(cell.getCellType()==0){  
					realValue=cell.getNumericCellValue()+"";
					if(realValue.endsWith(".0")){
						realValue=realValue.substring(0,realValue.length()-2);
					}else if(realValue.contains(".")){
						realValue=getRoundNum(realValue);
					}
				}else if(cell.getCellType()==2){
					realValue=cell.getNumericCellValue()+"";
					if(realValue.endsWith(".0")){
						realValue=realValue.substring(0,realValue.length()-2);
					}else if(realValue.contains(".")){
						realValue=getRoundNum(realValue);
					}
				}else{
					realValue="0";
				}
			}else{
				realValue="0";
			}
			realValue = realValue.replaceAll(",", "");
			ips.setString(postion, realValue);
		}
		
		private String getRoundNum(String inputNum){
			String outputNum=null;
			NumberFormat ddf1=NumberFormat.getNumberInstance() ;
			ddf1.setMaximumFractionDigits(3); 
			outputNum= ddf1.format(Double.parseDouble(inputNum)) ;
			return outputNum;
		}
	    
		private String getRoundNum(String value,int scale){
		   if ("".equals(value)|| "0".equals(value) || value==null)
				  return "0";
		   else{
		       BigDecimal   bd2   =   new   BigDecimal(value.trim());  
		       BigDecimal bd3=bd2.setScale(scale,   BigDecimal.ROUND_HALF_UP);
		     
		       return bd3.toString();
	       }		
		}	
		
		private void setIPSValueReq(PreparedStatement ips,Row rowp, int postion, int posValue, boolean type) throws SQLException{
			if(posValue>=0){
				Cell cell=null;
				if(type)
					cell= rowp.getCell(posValue);  
				else
					cell= rowp.getCell(posValue+1);  			           	
				setIPSValue1000(ips,cell,postion);
			}else{
				ips.setString(postion, "0");
			}
	   }	
		
	   /*
	    * 读取单元格式的值
	    */
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

		private void setIPSValue(PreparedStatement ips,Cell cell, int postion,int scale) throws SQLException{
			String realValue="";
			if(cell!=null){
				if(cell.getCellType()==1){    
					realValue=cell.getStringCellValue();
				}else if(cell.getCellType()==0){        
					realValue=cell.getNumericCellValue()+"";
					if(realValue.endsWith(".0")){
						realValue=realValue.substring(0,realValue.length()-2);
					}else if(realValue.contains(".")){
						realValue=getRoundNum(realValue,scale);
					}
				}else if(cell.getCellType()==2){
					realValue=cell.getNumericCellValue()+"";
					if(realValue.endsWith(".0")){
						realValue=realValue.substring(0,realValue.length()-2);
					}else if(realValue.contains(".")){
						realValue=getRoundNum(realValue,scale);
					}
				}else{
					realValue="";
					ips.setString(postion, null);
				}
			}else{
				realValue="";
			}
			ips.setString(postion, realValue);
		}
		
		
		
		private void setIPSValue(PreparedStatement ips,Cell cell, int postion) throws SQLException{
			String realValue="";
			if(cell!=null){
				if(cell.getCellType()==1){    
					realValue=cell.getStringCellValue();
				}else if(cell.getCellType()==0){        
					realValue=cell.getNumericCellValue()+"";
					if(realValue.endsWith(".0")){
						realValue=realValue.substring(0,realValue.length()-2);
					}else if(realValue.contains(".")){
						realValue=getRoundNum(realValue);
					}
				}else if(cell.getCellType()==2){
					realValue=cell.getNumericCellValue()+"";
					if(realValue.endsWith(".0")){
						realValue=realValue.substring(0,realValue.length()-2);
					}else if(realValue.contains(".")){
						realValue=getRoundNum(realValue);
					}
				}else{
					realValue="";
					ips.setString(postion, null);
				}
			}else{
				realValue="";
			}
			ips.setString(postion, realValue);
		}
	   
       private void SendError(String mes){
    	   OutRunTimeLog.showLog(mes);	
       }
		
		@SuppressWarnings("finally")
		public int setCommonExcelParam(PreparedStatement ips, String monitorId,Sheet sheet,int fDatePOS,String divCode) throws SQLException, ParseException{
			String strCell;
			int count=0;
			int rowNum = sheet.getLastRowNum(); 
			int colNum = sheet.getRow(0).getPhysicalNumberOfCells();
			try{
			    SimpleDateFormat formatter = new SimpleDateFormat(Keys.DATE_TIME_FORMAT2);
			    DateFormat df = new SimpleDateFormat(Keys.DATE_TIME_FORMAT3,Locale.US);
				Date dateTime=null;
				String DateStr;	 
				int pos;
				for (int i = 1; i <= rowNum; i++) {			
					Row row = sheet.getRow(i); 
				    for(int j=0; j< colNum; j++){		    	
				    	Cell cell = row.getCell(j);			    	
				    	if(fDatePOS > 0 && (fDatePOS-j)==1){ 
				    		DateStr = cell.getStringCellValue();
				    		 OutRunTimeLog.showLog(monitorId+"日期："+DateStr);
							if(! DateStr.equals("")){           //日期和时间
								if (DateStr.indexOf("下午") != -1){
									DateStr = DateStr.replaceFirst("下午 ","");
									  pos = DateStr.indexOf(":", 8);
									if(pos != -1){
										if("12".equals(DateStr.substring(pos-2, pos))){
											DateStr = DateStr.substring(0, pos-2)+"13"+DateStr.substring(pos);
										}else{
											DateStr = DateStr.substring(0, pos-2)+(Integer.parseInt(DateStr.substring(pos-2, pos))+12)+DateStr.substring(pos);
										}
									}
								}else if (DateStr.indexOf("CST") != -1){
									dateTime = df.parse(DateStr);
									DateStr = formatter.format(dateTime);
								}else{
									DateStr = DateStr.replaceFirst("上午 ","");
								}
								dateTime=formatter.parse(DateStr);
								ips.setDate(j+1, new java.sql.Date(dateTime.getTime()));
							}		
								
				    		/*
				    		if(cell == null ){						
				    		   ips.setDate(j+1, new java.sql.Date(new Date().getTime()));
							}else if(cell.getCellType()==1 || cell.getCellType()==2 ){
							
								
								DateStr = cell.getStringCellValue();
								if(DateStr.startsWith("'")){
									cell.setCellValue(cell.getStringCellValue().substring(2));
								}
								DateStr = cell.getStringCellValue();
								//System.out.println(j+"=="+(cell==null?"":cell.getStringCellValue()));
								if (DateStr.indexOf("CST") != -1){
									dateTime = df.parse(DateStr);
									DateStr = formatter.format(dateTime);						
									dateTime=formatter.parse(DateStr);
									ips.setDate(j+1, new java.sql.Date(dateTime.getTime()));
								}else if(cell.getStringCellValue().substring(4, 4).equals("-")){
								   ips.setDate(j+1, new java.sql.Date(new SimpleDateFormat("yyyy-MM-dd").parse(cell.getStringCellValue()).getTime()));
								}else if(cell.getStringCellValue().substring(4, 4).equals("/")){
									ips.setDate(j+1, new java.sql.Date(new SimpleDateFormat("yyyy/MM/dd").parse(cell.getStringCellValue()).getTime()));
								}else if(cell.getStringCellValue().substring(2, 2).equals("-")){
									ips.setDate(j+1, new java.sql.Date(new SimpleDateFormat("MM-dd-yyyy").parse(cell.getStringCellValue()).getTime()));
								}else if(cell.getStringCellValue().substring(2, 2).equals("/")){
									ips.setDate(j+1, new java.sql.Date(new SimpleDateFormat("MM/dd/yyyy").parse(cell.getStringCellValue()).getTime()));
								}else{
									ips.setDate(j+1, new java.sql.Date(new Date().getTime()));
								}							
							}else if(DateUtil.isCellDateFormatted(cell)){					
				               ips.setDate(j+1,new java.sql.Date(cell.getDateCellValue().getTime()));
				            } */
							
				    	}else{
				    		
				    		 strCell = getCellValue(cell); 	
				    		 if(strCell.endsWith(".0"))
				    			 strCell = strCell.substring(0, strCell.length()-2);
				    		 ips.setString(j+1, strCell);	
				    		
				    		 /**
				    		if(cell==null || cell.getCellType()==3){
				    			ips.setString(j+1, "");
				    		}else if(cell.getCellType()==0){
				    			ips.setString(j+1, cell==null?"0":((cell.getNumericCellValue()+"").endsWith(".0")?(cell.getNumericCellValue()+"").replace(".0", ""):(cell.getNumericCellValue()+"")));
				    		}else if(cell.getCellType()==2) {
				    			
				    			//cell.setCellFormula(cell.getCellFormula());			    			
				    	        strCell = getCellValue(cell); 			    		
				    			
				    			strCell = strCell.replace("'", "''");
				    			strCell = strCell.replace(",", "、");
				    			strCell = strCell.replace("&", "'|| CHR(38) ||'");
				    			
				    		    if (strCell.endsWith(".0") )
				    		       strCell = strCell.replace(".0", "");
				    		    
				    	        ips.setString(j+1, strCell==null?"":strCell.trim());	
				    	        
				    		}else{
				    			cell.setCellValue(cell.getStringCellValue().replace("'", "''"));
					    		cell.setCellValue(cell.getStringCellValue().replace(",", "、"));
					    		cell.setCellValue(cell.getStringCellValue().replace("&", "'|| CHR(38) ||'"));			
					        	ips.setString(j+1, cell==null?"":cell.getStringCellValue().trim());	
					        }	
					        */	    	  		    	   
				    	}
				    }
					if(fDatePOS == 0){			
						ips.setDate(colNum+1, new java.sql.Date(new Date().getTime()));
						ips.setString(colNum+2, divCode);
					}else{
						ips.setString(colNum+1, divCode);				
					}
				    ips.addBatch();
				    count = count + 1;
				}
			}catch (Exception e){
				e.printStackTrace();
				SendError(e.getMessage());			
			}finally{
				return count++;
			}
		}	

		@SuppressWarnings("finally")
		public int setFQEFQA396Param(PreparedStatement ips, String monitorId,Sheet sheet,int fDatePOS,String divCode) throws Exception{
			int count=0;
			try{
				Row row = null;
				Cell cell = null;			
				java.sql.Date date=null;			
			    String time,currValue;  
				int rows = sheet.getLastRowNum();
				//int Cols = sheet.getRow(9).getPhysicalNumberOfCells();		
		
				rows =rows -5 ;
				row = sheet.getRow(3);
				cell = row.getCell(14);	
		
				date = ParseUtil.CellValueToDate(cell);		
				
				cell = row.getCell(15);				
				time =getCellValue(cell);
				if(!"".equals(time) && time.length()>16) 
				   time =time.substring(11,16);		
				
				int j = 0,m=0;
				for (int i = 10; i <= rows; i++) {
					row = sheet.getRow(i);				    	
					cell = row.getCell(0);	
					currValue = getCellValue(row.getCell(0));
					
					if(cell != null  && !"".equals(currValue)) {  
						 j = 1;					 
						 ips.setDate(j++,date);	
						 ips.setString(j++, time);
						 ips.setString(j++, currValue);
						 
						 cell = row.getCell(3);	
						 currValue =getCellValue(cell);
						 if(currValue.endsWith(".0"))
							currValue = currValue.substring(0,currValue.length()-2);
						
						 ips.setString(j++, currValue);
						 
						
						 if(!"".equals(currValue) && currValue.length()>14)
							 currValue =currValue.substring(11,15);			
						
						 ips.setString(j++, currValue);
						
					
						ips.setString(j++, getCellValue(row.getCell(23))); 
						ips.setString(j++, getCellValue(row.getCell(2))+"--"+getCellValue(row.getCell(1)));
						
						for(int n=4;n<9;n++){
							ips.setString(j++, getCellValue(row.getCell(n)).replaceAll("\\+0",""));
						}
						
						ips.setString(j++, getCellValue(row.getCell(9))); 
						ips.setString(j++, getCellValue(row.getCell(9)));
						
						ips.setString(j++, getCellValue(row.getCell(10)));
						ips.setString(j++, getCellValue(row.getCell(10)));
						
						
						for(m=j;m<j+8;m++){ //10--8 删除   板凹(块数/点数) j=17
							ips.setString(m, getCellValue(row.getCell(m-3))); 
						}
					
						ips.setString(m, getCellValue(row.getCell(13)));
						ips.setString(m+1, getCellValue(row.getCell(23))); //25--23
						ips.setString(m+2, divCode); 
						
						ips.addBatch();
					    count = count + 1;				   
					 }else{
						 break;
					 }
					
				}				
		
			}catch (Exception e){
				e.printStackTrace();
				SendError(e.getMessage());			
			}finally{
				return count;
			}	
		}

		@SuppressWarnings("finally")
		public int setFQEROUD001Param(PreparedStatement ips, String monitorId,Sheet sheet,int fDatePOS,String divCode) throws Exception{
			int count=0;
			try{
				Row rowm = null;
				Cell cell = null;			
				Date dateTime=null;
				
				int rowNum = sheet.getLastRowNum();
				for (int i = 2; i <= rowNum; i++) {
					rowm = sheet.getRow(i);
					
					cell=rowm.getCell(2);     //型号
					if(cell==null)break;
					setIPSValue(ips,cell,4); 
					
					cell=rowm.getCell(28);     //日期
					if(cell!=null&&cell.getCellType()==0){          
						dateTime=cell.getDateCellValue();
						ips.setDate(1, new java.sql.Date(dateTime.getTime()));
					}else{
						ips.setString(1, null);
					}
					
					cell=rowm.getCell(29);     //时间
					if(cell!=null&&cell.getCellType()==0){          
						dateTime=cell.getDateCellValue();
						ips.setString(2, dateTime.getHours()+":"+dateTime.getMinutes());
					}else{
						ips.setString(2, null);
					}
					
					ips.setString(3, null);  //系列号
											
					cell=rowm.getCell(4);     //层次
					setIPSValue(ips,cell,5); 
					
					cell=rowm.getCell(17);     //阻抗值
					setIPSValue(ips,cell,6); 
					
					ips.setString(7, null); //备注		
					ips.setString(8, divCode); //备注
					
					ips.addBatch();
				    count = count + 1;
				}
			}catch (Exception e){
				e.printStackTrace();
				SendError(e.getMessage());
				
			}finally{
				return count;
			}	
		}
		

		@SuppressWarnings("finally")
		public int setDMELAM003AParam(PreparedStatement ips, String monitorId,Sheet sheet,int fDatePOS,String divCode) throws Exception{
			int count=0;
			try{
				Row rowm = null;
				Cell cell = null;	
				String PN=null;       //型号
				String LotNum=null;   //lot号
				String remark=null;
				String layerNum=null;
				Date dateTime=null;
				
				rowm = sheet.getRow(0);
				
				cell=rowm.getCell(1);			
				if(cell!=null&&cell.getCellType()==1){           
					PN=cell.getStringCellValue();
				}	
				cell=rowm.getCell(4);			
				if(cell!=null&&cell.getCellType()==1){           
					remark=cell.getStringCellValue();         //备注固定填写在第一行第五列
				}
				
				int rowNum = sheet.getLastRowNum();
				for (int i = 3; i <= rowNum; i=i+5) {
					ips.setString(3, PN);
					
					rowm = sheet.getRow(i);
					
					cell=rowm.getCell(3);				
					if(cell!=null&&cell.getCellType()==1){           //lot号
						LotNum=cell.getStringCellValue();
					}else{
						LotNum=null;
						break;
					}
					ips.setString(4, LotNum);
					
					ips.setString(5, layerNum);
					
					
					cell=rowm.getCell(4);
					if(cell!=null&&cell.getCellType()==0){           //日期和时间
						dateTime=cell.getDateCellValue();
						ips.setDate(1, new java.sql.Date(dateTime.getTime()));
						ips.setString(2, dateTime.getHours()+":"+dateTime.getMinutes());
					}else{
						ips.setString(1, null);
						ips.setString(2, null);
					}
					
					cell=rowm.getCell(44);
					setIPSValue(ips,cell,78);          //测试人
					
					cell=rowm.getCell(45);
					setIPSValue(ips,cell,79);           //评核
					
					cell=rowm.getCell(1);
					setIPSValue(ips,cell,6);           //A面BGA要求
					
					cell=rowm.getCell(6);
					setIPSValue(ips,cell,18);           //A面SMD要求
					
					cell=rowm.getCell(12);
					setIPSValue(ips,cell,30);           //A面01005长要求
					
					cell=rowm.getCell(18);
					setIPSValue(ips,cell,42);           //A面01005宽要求
					
					cell=rowm.getCell(23);
					setIPSValue(ips,cell,12);           //B面BGA要求
					
					cell=rowm.getCell(29);
					setIPSValue(ips,cell,24);           //B面SMD要求
					
					cell=rowm.getCell(34);
					setIPSValue(ips,cell,36);           //B面01005长要求
					
					cell=rowm.getCell(40);
					setIPSValue(ips,cell,48);           //B面01005宽要求
					
					ips.setString(80, remark);            //备注
					ips.setString(81, divCode);
					
					
					for(int j=0; j<5; j++){        		
						Row rown = sheet.getRow(i+j);
						if(rown!=null){
							cell=rown.getCell(3);          //lot号
							if(cell!=null&&cell.getCellType()==1){ 
								if(LotNum.equals(cell.getStringCellValue())){
									cell=rown.getCell(2);              //A面BGA实测值		
									setIPSValue(ips,cell,7+j);
									cell=rown.getCell(7);              //A面SMD实测值		
									setIPSValue(ips,cell,19+j);
									cell=rown.getCell(13);              //A面01005长实测值		
									setIPSValue(ips,cell,31+j);
									cell=rown.getCell(19);              //A面01005宽实测值		
									setIPSValue(ips,cell,43+j);
									
									cell=rown.getCell(24);              //B面BGA实测值		
									setIPSValue(ips,cell,13+j);
									cell=rown.getCell(30);              //B面SMD实测值		
									setIPSValue(ips,cell,25+j);
									cell=rown.getCell(35);              //B面01005长实测值		
									setIPSValue(ips,cell,37+j);
									cell=rown.getCell(41);              //B面01005宽实测值		
									setIPSValue(ips,cell,49+j);
									
								}else{
									for(int k=j; k<5;k++){
										ips.setString(7+k, null);
										ips.setString(19+k, null);
										ips.setString(31+k, null);
										ips.setString(43+k, null);
										ips.setString(13+k, null);
										ips.setString(25+k, null);
										ips.setString(37+k, null);
										ips.setString(49+k, null);
									}
									i=i-(5-j);
									break;
								}
							}else{
								for(int k=j; k<5;k++){
									ips.setString(7+k, null);
									ips.setString(19+k, null);
									ips.setString(31+k, null);
									ips.setString(43+k, null);
									ips.setString(13+k, null);
									ips.setString(25+k, null);
									ips.setString(37+k, null);
									ips.setString(49+k, null);
								}
								i=i-(5-j);
								break;
							}						
						}else{
							for(int k=j; k<5;k++){
								ips.setString(7+k, null);
								ips.setString(19+k, null);
								ips.setString(31+k, null);
								ips.setString(43+k, null);
								ips.setString(13+k, null);
								ips.setString(25+k, null);
								ips.setString(37+k, null);
								ips.setString(49+k, null);
							}
							break;
						}					
					}
					
					for(int j=54; j<78; j++){             //0201数据暂时不确定
						ips.setString(j, null);
					}
					
					ips.addBatch();
				    count = count + 1;
				}
			}catch (Exception e){
				e.printStackTrace();
				SendError(e.getMessage());
				//System.out.println(e.getMessage());
			}finally{
				return count;
			}	
		}
		
		@SuppressWarnings("finally")
		public int setFQEPQA366AParam(PreparedStatement ips, String monitorId,Sheet sheet,int fDatePOS,String divCode) throws Exception{
			int count=0;
			try{
				Row rowm = null;
				Cell cell = null;	
				String PN=null;       //型号
				String LotNum=null;   //lot号
				String remark=null;   //备注
				Date dateTime=null;
				final int colInterval=6;
							 
				rowm = sheet.getRow(0);
				cell=rowm.getCell(1);
				
				if(cell!=null&&cell.getCellType()==1){           
					PN=cell.getStringCellValue();
				}	
				
				cell=rowm.getCell(4);			
				if(cell!=null&&cell.getCellType()==1){           
					remark=cell.getStringCellValue();        //备注固定填写在第一行第五列
				}	
				int rowNum = sheet.getLastRowNum();
				for (int i = 3; i <= rowNum; i=i+5) {
					ips.setString(3, PN);
					
					rowm = sheet.getRow(i);
					
					cell=rowm.getCell(4);				
					if(cell!=null&&cell.getCellType()==1){           //lot号
						LotNum=cell.getStringCellValue();
					}else{
						LotNum=null;
						break;
					}
					ips.setString(4, LotNum);
									
					cell=rowm.getCell(5);          //时间	
					if(cell!=null&&cell.getCellType()==0){           //日期和时间
						dateTime=cell.getDateCellValue();
						ips.setDate(1, new java.sql.Date(dateTime.getTime()));
						ips.setString(2, dateTime.getHours()+":"+dateTime.getMinutes());
					}else{
						ips.setString(1, null);
						ips.setString(2, null);
					}
					
					cell=rowm.getCell(36);
					setIPSValue(ips,cell,41);          //测试人
					
					cell=rowm.getCell(37);
					setIPSValue(ips,cell,42);           //评核
					for (int classifyNum=0; classifyNum<3; classifyNum++){
						cell=rowm.getCell(1+colInterval*classifyNum);					//A面单元内线宽要求、SMD直径要求、BGA直径要求
						setIPSValue(ips,cell,5+colInterval*classifyNum*2);
						cell=rowm.getCell(19+colInterval*classifyNum);					//B面单元内线宽要求、SMD直径要求、BGA直径要求
						setIPSValue(ips,cell,11+colInterval*classifyNum*2);	
						
						for(int j=0; j<5; j++){        			
							Row rown = sheet.getRow(i+j);
							if(rown!=null){
								cell=rown.getCell(4);	
								if(cell!=null&&cell.getCellType()==1){           
									if(LotNum.equals(cell.getStringCellValue())){
										cell=rown.getCell(2+colInterval*classifyNum);          //A面单元内线宽	、SMD直径、BGA直径
										setIPSValue(ips,cell,6+colInterval*classifyNum*2+j);
										cell=rown.getCell(20+colInterval*classifyNum);          //B面单元内线宽	、SMD直径、BGA直径	
										setIPSValue(ips,cell,12+colInterval*classifyNum*2+j);
									}else{
										for(int k=j; k<5;k++){
											ips.setString(6+colInterval*classifyNum*2+k, null);
											ips.setString(12+colInterval*classifyNum*2+k, null);
										}
										if(classifyNum==2) i=i-(5-j);
										break;
									}
								}else{
									for(int k=j; k<5;k++){
										ips.setString(6+colInterval*classifyNum*2+k, null);
										ips.setString(12+colInterval*classifyNum*2+k, null);
									}
									if(classifyNum==2) i=i-(5-j);
									break;
								}
							}else{
								for(int k=j; k<5;k++){
									ips.setString(6+colInterval*classifyNum*2+k, null);
									ips.setString(12+colInterval*classifyNum*2+k, null);								
								}
								if(classifyNum==2) i=i-(5-j);
								break;
							}
						}					
					}
					
					ips.setString(43, remark);            //备注
					ips.setString(44, divCode);
					ips.addBatch();
				    count = count + 1;
				}
			}catch (Exception e){
				SendError(e.getMessage());
				e.printStackTrace();
				//System.out.println(e.getMessage());
			}finally{
				return count;
			}	
		}

		
		@SuppressWarnings("finally")
		public int ImpQCIDF01Data(PreparedStatement ips, String monitorId,Sheet sheet,int fDatePOS,String divCode) throws Exception{
			int count=0;
			try{
				Row rowm = null,rowHead;
				Cell cell = null;	
				Date dateTime=null;
				String DateStr;	 
			    SimpleDateFormat formatter = new SimpleDateFormat(Keys.DATE_TIME_FORMAT2);
			    DateFormat df = new SimpleDateFormat(Keys.DATE_TIME_FORMAT3,Locale.US);
			    int pos=0;
				rowm = sheet.getRow(0);				
				String PN = getCellValue(rowm.getCell(1));
			
				rowm = sheet.getRow(3);	
				String LayerNum = getCellValue(rowm.getCell(1));
			
				rowm = sheet.getRow(5);			
				String URL =  getCellValue(rowm.getCell(3));		
				String LRL =  getCellValue(rowm.getCell(4));
				if(URL.endsWith(".0")){
					URL = URL.substring(0, URL.length()-2) ;
				}
				if(LRL.endsWith(".0")){
					LRL = LRL.substring(0, LRL.length()-2) ;
				}
				String Lot;
				int rowNum = sheet.getLastRowNum();
				int firstRow = 12;
				rowHead = sheet.getRow(11);	
				Lot = getCellValue(rowHead.getCell(0));
				if("CPK".equals(Lot.toUpperCase())){
					firstRow = 14;
				}else if("LOT".equals(Lot.toUpperCase())){
					firstRow =12;
				}
		
				if( firstRow >= rowNum){
					return -1;
				}
				for (int i =firstRow; i <= rowNum; i=i+5) {
					rowm = sheet.getRow(i);	
					Lot = getCellValue(rowm.getCell(0));
					cell=rowm.getCell(4);				
					
					if(cell ==null ){
						ips.setDate(1, new java.sql.Date(new Date().getTime()));
						ips.setString(2, "晚班");					
					}else{ 
						DateStr = getCellValue(cell);					
						if(! DateStr.equals("")){           //日期和时间
							if (DateStr.indexOf("下午") != -1){
								DateStr = DateStr.replaceFirst("下午 ","");
								  pos = DateStr.indexOf(":", 8);
								if(pos != -1){
									if("12".equals(DateStr.substring(pos-2, pos))){
										DateStr = DateStr.substring(0, pos-2)+"13"+DateStr.substring(pos);
									}else{
										DateStr = DateStr.substring(0, pos-2)+(Integer.parseInt(DateStr.substring(pos-2, pos))+12)+DateStr.substring(pos);
									}
								}
							}else if (DateStr.indexOf("CST") != -1){
								dateTime = df.parse(DateStr);
								DateStr = formatter.format(dateTime);
							}else{
								DateStr = DateStr.replaceFirst("上午 ","");
							}
							dateTime=formatter.parse(DateStr);
							ips.setDate(1, new java.sql.Date(dateTime.getTime()));
							ips.setString(2, dateTime.getHours()+":"+dateTime.getMinutes());
						}else{					
							ips.setDate(1, new java.sql.Date(new Date().getTime()));
							ips.setString(2, "晚班");
						}	
					}
					
					ips.setString(3, PN);
					ips.setString(4, Lot);
					ips.setString(5, LayerNum);
					ips.setString(6, LRL);
					ips.setString(7, URL);	
					
					for(int j=0; j<5; j++){         //实测值				
						Row rown = sheet.getRow(i+j);
						Cell subcell = rown.getCell(3);
						if(Lot.equals(getCellValue(rown.getCell(0)))){
						   setIPSValue(ips,subcell,8+j,2);			
						}else{
						   ips.setString(8+j, null);	
						}
					}
					       
					ips.setString(13, getCellValue(rowm.getCell(5)));	//  //测试人
					ips.setString(14, getCellValue(rowm.getCell(6)));	//评核				
					ips.setString(15, getCellValue(rowm.getCell(7)));      //备注
					
					ips.setString(16, divCode);
					ips.addBatch();
					
				    count = count + 1;
				}
			}catch (Exception e){
				SendError(e.getMessage());			
			}finally{
				return count;
			}	
		}
		
		@SuppressWarnings("finally")
		public int setFQEPQA410AParam(PreparedStatement ips, String monitorId,Sheet sheet,int fDatePOS,String divCode) throws Exception{
			int count=0;
			try{
				Row rowm = null;
				Cell cell = null;	
				String PN=null;       //型号
				String LayerNum=null; //层次
				String LotNum=null;   //lot号
				String remark=null;   //备注
				Date dateTime=null;
							 
				rowm = sheet.getRow(0);
				cell=rowm.getCell(1);
				if(cell!=null&&cell.getCellType()==1){           
					PN=cell.getStringCellValue();
				}
				cell=rowm.getCell(4);
				if(cell!=null&&cell.getCellType()==1){          
					LayerNum=cell.getStringCellValue();      //层次固定填写在第一行第五列
				}		
				cell=rowm.getCell(6);			
				if(cell!=null&&cell.getCellType()==1){           
					remark=cell.getStringCellValue();       //备注固定填写在第一行第七列
				}
				
				int rowNum = sheet.getLastRowNum();
				for (int i = 3; i <= rowNum; i=i+5) {
					ips.setString(3, PN);
					ips.setString(5, LayerNum);
					
					rowm = sheet.getRow(i);
					cell=rowm.getCell(1);	        //要求
					setIPSValue(ips,cell,6);

					cell=rowm.getCell(4);
					if(cell!=null&&cell.getCellType()==1){           //lot号
						LotNum=cell.getStringCellValue();
					}else{
						LotNum=null;
						break;
					}
					ips.setString(4, LotNum);
					
					cell=rowm.getCell(5);
					if(cell!=null&&cell.getCellType()==0){           //日期和时间
						dateTime=cell.getDateCellValue();
						ips.setDate(1, new java.sql.Date(dateTime.getTime()));
						ips.setString(2, dateTime.getHours()+":"+dateTime.getMinutes());
					}else{
						ips.setString(1, null);
						ips.setString(2, null);
					}
					
					cell=rowm.getCell(7);
					setIPSValue(ips,cell,12);          //测试人
					
					cell=rowm.getCell(8);
					setIPSValue(ips,cell,13);          //评核
					
					for(int j=0; j<5; j++){         //实测值				
						Row rown = sheet.getRow(i+j);
						if(rown!=null){
							cell=rown.getCell(4);	
							if(cell!=null&&cell.getCellType()==1){           
								if(LotNum.equals(cell.getStringCellValue())){
									cell=rown.getCell(3);
									setIPSValue(ips,cell,7+j);	
								}else{
									for(int k=j; k<5;k++){
										ips.setString(7+k, null);
									}
									i=i-(5-j);
									break;
								}
							}else{
								for(int k=j; k<5;k++){
									ips.setString(7+k, null);
								}
								i=i-(5-j);
								break;
							}						
						}else{
							for(int k=j; k<5;k++){
								ips.setString(7+k, null);
							}
							break;
						}
					}
					ips.setString(14, remark);       //备注
					ips.setString(15, divCode);
					ips.addBatch();
				    count = count + 1;
				}
			}catch (Exception e){
				SendError(e.getMessage());
				e.printStackTrace();		
			}finally{
				return count;
			}	
		}
		
		@SuppressWarnings("finally")
		public int setFQEPQA410AParam2(PreparedStatement ips, String monitorId,Sheet sheet,int fDatePOS,String divCode) throws Exception{
			int count=0;
			try{
				Row rowm = null;
				Cell cell = null;	
				String PN=null;       //型号
				String LayerNum1=null; //层次
				String LayerNum2=null; //层次
				String LotNum=null;   //lot号
				String remark1=null;   //备注1
				String remark2=null;   //备注2
				String assessment=null; //评核
				String tester=null;   //测试人
				
				Date dateTime=null;     //日期时间
							 
				rowm = sheet.getRow(0);
				cell=rowm.getCell(1);
				if(cell!=null&&cell.getCellType()==1){           
					PN=cell.getStringCellValue();                 //型号
					if(PN.contains("-")){
						LayerNum1="L"+(PN.substring(PN.length()-4, PN.length()-3).equals("0")?PN.substring(PN.length()-3, PN.length()-2):PN.substring(PN.length()-4, PN.length()-2));         //层次				
						LayerNum2="L"+(PN.substring(PN.length()-2,PN.length()-1).equals("0")?PN.substring(PN.length()-1):PN.substring(PN.length()-2));         //层次	
						PN=PN.substring(0,PN.lastIndexOf('-'));					
					}

				}
				
				int rowNum = sheet.getLastRowNum();
				for (int i = 3; i <= rowNum; i=i+5) {
					//日期 1 ： 时间2： 型号3：流程卡号4：层次5：宽要求6：unit1 7：unit2 8： unit3 9： unit4 10： init5 11： 测试人12： 评核13： 备注14
					//层次1
					ips.setString(3, PN);
					ips.setString(5, LayerNum1);
//					ips.setString(12, tester);
//					ips.setString(13, assessment);					
					ips.setString(14, remark1);       //备注
					
					ips.setString(15, "");       //Add 2016-9-14
					ips.setString(16, "");       //Add 2016-9-14
					ips.setString(17, "");       //Add 2016-9-14
					
					ips.setString(18, divCode);
					
					rowm = sheet.getRow(i);
					cell=rowm.getCell(1);	        //要求
					setIPSValue(ips,cell,6);

					cell=rowm.getCell(4);
					if(cell!=null&&cell.getCellType()==1){           //lot号
						LotNum=cell.getStringCellValue();
					}else{
						LotNum=null;
						break;
					}
					ips.setString(4, LotNum);
					
					cell=rowm.getCell(5);
					if(cell!=null&&cell.getCellType()==0){           //日期和时间
						dateTime=cell.getDateCellValue();
						ips.setDate(1, new java.sql.Date(dateTime.getTime()));
						ips.setString(2, dateTime.getHours()+":"+dateTime.getMinutes());
					}else{
						ips.setString(1, null);
						ips.setString(2, null);
					}
							
					if(rowm.getCell(6)!=null&&rowm.getCell(12)!=null){
						cell=rowm.getCell(12);	             //区分不同格式表格评核人填写的位置不同
					}else{
						cell=rowm.getCell(13);
					}					       
					if(cell!=null&&cell.getCellType()==1){           
						assessment=cell.getStringCellValue().toUpperCase().substring(0, 3);
						if("ACC".equals(assessment)){
							tester=cell.getStringCellValue().replace("ACC", "");
						}else if("REJ".equals(assessment)){
							tester=cell.getStringCellValue().replace("REJ", "");
						}else{
							tester=cell.getStringCellValue();
							assessment=null;
						}
					}else{
						tester=null;
						assessment=null;
					}
					ips.setString(12, tester);     //测试人
					ips.setString(13, assessment);   //评核
					
					
					for(int j=0; j<5; j++){         //实测值				
						Row rown = sheet.getRow(i+j);
						if(rown!=null){
							cell=rown.getCell(4);	
							if(cell!=null&&cell.getCellType()==1){           
								if(LotNum.equals(cell.getStringCellValue())){
									cell=rown.getCell(2);
									//setIPSValue1000(ips,cell,7+j);	
									setIPSValueForInner(ips,cell,7+j);	
								}else{
									for(int k=j; k<5;k++){
										ips.setString(7+k, null);
									}
									i=i-(5-j);
									break;
								}
							}else{
								for(int k=j; k<5;k++){
									ips.setString(7+k, null);
								}
								i=i-(5-j);
								break;
							}						
						}else{
							for(int k=j; k<5;k++){
								ips.setString(7+k, null);
							}
							break;
						}
					}
					ips.addBatch();
				    count = count + 1;    	    			    
				}
				for (int i = 3; i <= rowNum; i=i+5) {
					//日期 1 ： 时间2： 型号3：流程卡号4：层次5：宽要求6：unit1 7：unit2 8： unit3 9： unit4 10： init5 11： 测试人12： 评核13： 备注14
				    
				  //层次2 
					int cellType=0;
					rowm = sheet.getRow(i);
					if(rowm.getCell(6)!=null){
						cellType=-1;                //区分不同格式表格
					}
					
					cell=rowm.getCell(11+cellType);
					if(cell!=null&&cell.getCellType()==1){           //lot号
						LotNum=cell.getStringCellValue();
					}else{
						LotNum=null;
						break;
					}
					ips.setString(4, LotNum);
					
					
					if(rowm.getCell(6)!=null&&rowm.getCell(12)!=null){
						cell=rowm.getCell(12);	             //区分不同格式表格评核人填写的位置不同
					}else{
						cell=rowm.getCell(13);
					}	
					if(cell!=null&&cell.getCellType()==1){           
						assessment=cell.getStringCellValue().toUpperCase().substring(0, 3);
						if("ACC".equals(assessment)){
							tester=cell.getStringCellValue().replace("ACC", "");
						}else if("REJ".equals(assessment)){
							tester=cell.getStringCellValue().replace("REJ", "");
						}else{
							tester=cell.getStringCellValue();
							assessment=null;
						}
					}else{
						tester=null;
						assessment=null;
					}
					ips.setString(3, PN);
					ips.setString(5, LayerNum2);
					ips.setString(12, tester);
					ips.setString(13, assessment);
					ips.setString(14, remark2);       //备注
					
					ips.setString(15, "");       //Add 2016-9-14
					ips.setString(16, "");       //Add 2016-9-14
					ips.setString(17, "");       //Add 2016-9-14
					
					ips.setString(18, divCode);
					
					rowm = sheet.getRow(i);
					cell=rowm.getCell(8+cellType);	        //要求
					setIPSValue(ips,cell,6);


					
					cell=rowm.getCell(12+cellType);
					if(cell!=null&&cell.getCellType()==0){           //日期和时间
						dateTime=cell.getDateCellValue();
						ips.setDate(1, new java.sql.Date(dateTime.getTime()));
						ips.setString(2, dateTime.getHours()+":"+dateTime.getMinutes());
					}else{
						ips.setString(1, null);
						ips.setString(2, null);
					}
									
					for(int j=0; j<5; j++){         //实测值				
						Row rown = sheet.getRow(i+j);
						if(rown!=null){
							cell=rown.getCell(11+cellType);	
							if(cell!=null&&cell.getCellType()==1){           
								if(LotNum.equals(cell.getStringCellValue())){
									cell=rown.getCell(9+cellType);
									//setIPSValue1000(ips,cell,7+j);	
									setIPSValueForInner(ips,cell,7+j);	
								}else{
									for(int k=j; k<5;k++){
										ips.setString(7+k, null);
									}
									i=i-(5-j);
									break;
								}
							}else{
								for(int k=j; k<5;k++){
									ips.setString(7+k, null);
								}
								i=i-(5-j);
								break;
							}						
						}else{
							for(int k=j; k<5;k++){
								ips.setString(7+k, null);
							}
							break;
						}
					}
					ips.addBatch();
				    count = count + 1;			    			    
				}
			}catch (Exception e){
				e.printStackTrace();
				SendError(e.getMessage());
			}finally{
				return count;
			}	
		}
		
		@SuppressWarnings("finally")
		public int setFQEPQA366AParam3(PreparedStatement ips, String monitorId,Sheet sheet,int fDatePOS,String divCode) throws Exception{
			int count=0;
			try{
				Row rowm = null;
				Row rowPos = null;
				Cell cell = null;	
				String PN=null;       //型号
				String LotNum=null;   //lot号
				String remark=null;   //备注
				String LayerNum=null;
				String layerNumA="99";
				String layerNumB="99";
				Date dateTime=null;
				int[] typePos={-1,-1,-1,-1,-1,-1};   //position of 线宽  SMD BGA for side A and side B
		
				String dataPos= null;
				
				rowm = sheet.getRow(0);
				rowPos = sheet.getRow(1);
				try{
					for(int p=0; p<41;p++){				
						cell =  rowPos.getCell(p);	
					    if(cell != null){
					    	count=count+1;				    	
					    }
					}				
					if(count!=6){
					   rowPos = sheet.getRow(2);			  
					}
				}catch(Exception e){				
					rowPos = sheet.getRow(2);	
				}
			
				
				count=0;
				
				cell=rowm.getCell(1);
				
				if(cell!=null ){           
					PN=getCellValue(cell);				
					if(PN.contains("-")){
						LayerNum="L"+(PN.substring(PN.length()-4, PN.length()-3).equals("0")?PN.substring(PN.length()-3, PN.length()-2):PN.substring(PN.length()-4, PN.length()-2))       				
						          +"/"+(PN.substring(PN.length()-2,PN.length()-1).equals("0")?PN.substring(PN.length()-1):PN.substring(PN.length()-2));    //层次
						layerNumA="L"+(PN.substring(PN.length()-4, PN.length()-3).equals("0")?PN.substring(PN.length()-3, PN.length()-2):PN.substring(PN.length()-4, PN.length()-2));
						layerNumB="L"+(PN.substring(PN.length()-2,PN.length()-1).equals("0")?PN.substring(PN.length()-1):PN.substring(PN.length()-2)); 
						PN=PN.substring(0,PN.lastIndexOf('-'));				
					}
				}	
				
				cell=rowm.getCell(4);			
				if(cell!=null ){           
					remark=getCellValue(cell);        //备注固定填写在第一行第五列
				}	
				remark=LayerNum;           //层次填写在备注栏位
				
				
				for(int pos=0; pos<41;pos++){
					cell=rowPos.getCell(pos);
					if(cell!=null&&cell.getCellType()==1){           
						dataPos=cell.getStringCellValue(); 
						if(dataPos.toUpperCase().contains(layerNumA)&&!dataPos.toUpperCase().contains(layerNumB)){
							if(dataPos.toUpperCase().contains("BGA")){
								typePos[2]=pos;
							}else if(dataPos.toUpperCase().contains("SMD")){
								typePos[1]=pos;
							}else{
								typePos[0]=pos;
							}
						}else{
							if(dataPos.toUpperCase().contains("BGA")){
								typePos[5]=pos;
							}else if(dataPos.toUpperCase().contains("SMD")){
								typePos[4]=pos;
							}else{
								typePos[3]=pos;
							}
						}
					}
				}
				
				int startPos=typePos[0];
				
				int rowNum = sheet.getLastRowNum();
				int Interval=0;
				for (int i = 3; i <= rowNum; i=i+5) {
					rowm = sheet.getRow(i);					
					for(int arrayLen=0 ;arrayLen < typePos.length; arrayLen++){
						if(typePos[arrayLen]>0 ){
							startPos=typePos[arrayLen];
							cell=rowm.getCell(startPos+3);	
							LotNum = getCellValue(cell); 
							
							if(!"".equals(LotNum)){
								break; //退出循环
							}
						}
					}	
									
					cell=rowm.getCell(startPos+4);          //时间
					
					if(cell!=null&&cell.getCellType()==0){           //日期和时间
						dateTime=cell.getDateCellValue();
						ips.setDate(1, new java.sql.Date(dateTime.getTime()));
						ips.setString(2, dateTime.getHours()+":"+dateTime.getMinutes());
					}else{
						ips.setString(1, null);
						ips.setString(2, null);
					}
					
					ips.setString(3, PN);
					ips.setString(4, LotNum);				
					String assessment=null;
					String tester=null;
					cell=rowm.getCell(41);      
					if(cell!=null&&cell.getCellType()==1){           
						assessment=cell.getStringCellValue().toUpperCase().substring(0, 3);
						if("ACC".equals(assessment)){
							tester=cell.getStringCellValue().replace("ACC", "");
						}else if("REJ".equals(assessment)){
							tester=cell.getStringCellValue().replace("REJ", "");
						}else{
							tester=cell.getStringCellValue();
							assessment=null;
						}
					}else{
						tester=null;
						assessment=null;
					}
					ips.setString(41, tester);     //测试人
					ips.setString(42, assessment);   //评核
					
					setIPSValueReq(ips,rowm,5, typePos[0],true);   //A面线宽要求
			         
					setIPSValueReq(ips,rowm,17, typePos[1],true);   //A面SMD要求
				           
					setIPSValueReq(ips,rowm,29, typePos[2],true);   //A面BGA要求
	        
					setIPSValueReq(ips,rowm,11, typePos[3],true);   //B面线宽要求
	    
					setIPSValueReq(ips,rowm,23, typePos[4],true);   //B面SMD要求
	    
					setIPSValueReq(ips,rowm,35, typePos[5],true);   //B面BGA要求

						
					
					for(int j=0; j<5; j++){        			
						Row rown = sheet.getRow(i+j);					
						if(rown!=null){						
							for(int F=0;F<typePos.length;F++){							
								switch (F) {
								case 0:
									   Interval=6;   //A面线宽要求
									   break;
								case 1:
									   Interval=18;  //A面SMD实测值
									   break;
								case 2:
									   Interval=30;  //A面BGA要求	
									   break;
								case 3:
									   Interval=12;   //B面线宽要求	
									   break;								   
								case 4:
									   Interval=24;   //B面SMD实测值
									   break;
								case 5:
									   Interval=36;   //B面BGA要求
									   break;								   
								default:
									   Interval=6;
									break;
								}
								
								cell=rown.getCell(typePos[F]+3);	
								if(cell!=null&&cell.getCellType()==1){           
									if(LotNum.equals(cell.getStringCellValue())){
										setIPSValueReq(ips,rown,Interval+j, typePos[F],false);      
									}else{
										ips.setString(Interval+j, "0");	
									}
									
								}else{
									ips.setString(Interval+j, "0");
								}	
														
							}			
							
							/*
							if(cell!=null&&cell.getCellType()==1){           
								if(LotNum.equals(cell.getStringCellValue())){
										
									setIPSValueReq(ips,rown,6+j, typePos[0],false);      //A面线宽要求	
										
									setIPSValueReq(ips,rown,18+j, typePos[1],false);     //A面SMD实测值

									setIPSValueReq(ips,rown,30+j, typePos[2],false);     //A面BGA要求	

									setIPSValueReq(ips,rown,12+j, typePos[3],false);      //B面线宽要求	

									setIPSValueReq(ips,rown,24+j, typePos[4],false);     //B面SMD实测值

									setIPSValueReq(ips,rown,36+j, typePos[5],false);        //B面BGA要求	

								}else{
									for(int k=j; k<5;k++){
										ips.setString(6+k, "0");
										ips.setString(18+k, "0");
										ips.setString(30+k, "0");
										ips.setString(12+k, "0");
										ips.setString(24+k, "0");
										ips.setString(36+k, "0");
									}
									i=i-(5-j);
									break;
								}
								
							}else{
								for(int k=j; k<5;k++){
									ips.setString(6+k, "0");
									ips.setString(18+k, "0");
									ips.setString(30+k, "0");
									ips.setString(12+k, "0");
									ips.setString(24+k, "0");
									ips.setString(36+k, "0");
								}
								i=i-(5-j);
								break;
							}	
							
							*/
						}else{
							for(int k=j; k<5;k++){
								ips.setString(6+k, "0");
								ips.setString(18+k, "0");
								ips.setString(30+k, "0");
								ips.setString(12+k, "0");
								ips.setString(24+k, "0");
								ips.setString(36+k, "0");
							}
							break;
						}					
					}					
					
					ips.setString(43, remark);            //备注
					ips.setString(44, divCode);
					ips.addBatch();
				    count = count + 1;
				}
			}catch (Exception e){
				e.printStackTrace();
				//System.out.println(e.getMessage());
			}finally{
				return count;
			}	
		}
		
		@SuppressWarnings("finally")
		public int setDMELAM003AParam3(PreparedStatement ips, String monitorId,Sheet sheet,int fDatePOS,String divCode) throws Exception{
			int count=0;
			try{
				Row rowm = null;
				Row rowPos = null;
				Cell cell = null;	
				String PN=null;       //型号
				String LotNum=null;   //lot号
				String remark=null;
				String layerNum="99";
				String layerNumA="99";
				String layerNumB="99";
				Date dateTime=null;
				int[] typePos={-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1};   //position of BGA  SMD   01005L    01005W    0201L   0201W  for side A and side B
				String dataPos= null;
				
				rowm = sheet.getRow(0);
				rowPos = sheet.getRow(2);
				
				cell=rowm.getCell(1);			
				if(cell!=null&&cell.getCellType()==1){           
					PN=cell.getStringCellValue();
					if(PN.contains("-")){
						layerNum="L"+(PN.substring(PN.length()-4, PN.length()-3).equals("0")?PN.substring(PN.length()-3, PN.length()-2):PN.substring(PN.length()-4, PN.length()-2))       				
								 +"/"+(PN.substring(PN.length()-2,PN.length()-1).equals("0")?PN.substring(PN.length()-1):PN.substring(PN.length()-2));    //层次
						layerNumA="L"+(PN.substring(PN.length()-4, PN.length()-3).equals("0")?PN.substring(PN.length()-3, PN.length()-2):PN.substring(PN.length()-4, PN.length()-2)) ;      				
						layerNumB="L"+(PN.substring(PN.length()-2,PN.length()-1).equals("0")?PN.substring(PN.length()-1):PN.substring(PN.length()-2));    //层次
						PN=PN.substring(0,PN.lastIndexOf('-'));  					
					}
				}	
				cell=rowm.getCell(4);			
				if(cell!=null&&cell.getCellType()==1){           
					remark=cell.getStringCellValue();         //备注固定填写在第一行第五列
				}
				
				for(int pos=0; pos<82;pos++){
					cell=rowPos.getCell(pos);
					if(cell!=null&&cell.getCellType()==1){           
						dataPos=cell.getStringCellValue(); 
						if(dataPos.toUpperCase().contains("BGA")&&dataPos.toUpperCase().contains(layerNumA)&&!dataPos.toUpperCase().contains(layerNumB)){    //L1/10
							typePos[0]=pos;
						}else if(dataPos.toUpperCase().contains("SMD")&&dataPos.toUpperCase().contains(layerNumA)&&!dataPos.toUpperCase().contains(layerNumB)){
							typePos[1]=pos;
						}else if(dataPos.toUpperCase().contains("01005")&&dataPos.toUpperCase().contains(layerNumA)&&!dataPos.toUpperCase().contains(layerNumB)){
							if(typePos[2]==-1){
								typePos[2]=pos;
							}else{
								typePos[3]=pos;
							}						
						}else if(dataPos.toUpperCase().contains("0201")&&dataPos.toUpperCase().contains(layerNumA)&&!dataPos.toUpperCase().contains(layerNumB)){
							if(typePos[4]==-1){
								typePos[4]=pos;
							}else{
								typePos[5]=pos;
							}						
						}else if(dataPos.toUpperCase().contains("BGA")&&dataPos.toUpperCase().contains(layerNumB)){
							typePos[6]=pos;
						}else if(dataPos.toUpperCase().contains("SMD")&&dataPos.toUpperCase().contains(layerNumB)){
							typePos[7]=pos;
						}else if(dataPos.toUpperCase().contains("01005")&&dataPos.toUpperCase().contains(layerNumB)){
							if(typePos[8]==-1){
								typePos[8]=pos;
							}else{
								typePos[9]=pos;
							}					
						}else if(dataPos.toUpperCase().contains("0201")&&dataPos.toUpperCase().contains(layerNumB)){
							if(typePos[10]==-1){
								typePos[10]=pos;
							}else{
								typePos[11]=pos;
							}
						}
					}
				}
				
				int rowNum = sheet.getLastRowNum();
				int startPos=typePos[0];
				int Interval=0;

				for (int i = 3; i <= rowNum; i=i+5) {	
					rowm = sheet.getRow(i);  						
					
					/**************
					if(cell!=null&&cell.getCellType()==1){           //lot号
						LotNum=cell.getStringCellValue();
					}else{
						LotNum=null;
						continue;
					} */
									
					for(int arrayLen=0 ;arrayLen < typePos.length; arrayLen++){
						if(typePos[arrayLen]>0 ){
							startPos=typePos[arrayLen];
							cell=rowm.getCell(startPos+3);	
							LotNum = getCellValue(cell); 
							
							if(!"".equals(LotNum)){
								break; //退出循环
							}
						}
					}	
									
					cell=rowm.getCell(startPos+4);          //时间
					if(cell!=null&&cell.getCellType()==0){           //日期和时间
						dateTime=cell.getDateCellValue();
						ips.setDate(1, new java.sql.Date(dateTime.getTime()));
						ips.setString(2, dateTime.getHours()+":"+dateTime.getMinutes());
					}else{
						ips.setString(1, null);
						ips.setString(2, null);
					}
					
					ips.setString(3, PN);
					ips.setString(4, LotNum);
					
					ips.setString(5, layerNum);
					
					cell=rowm.getCell(82);  
					String assessment=null;
					String tester=null;      
					if(cell!=null&&cell.getCellType()==1){           
						assessment=cell.getStringCellValue().toUpperCase().substring(0, 3);
						if("ACC".equals(assessment)){
							tester=cell.getStringCellValue().replace("ACC", "");
						}else if("REJ".equals(assessment)){
							tester=cell.getStringCellValue().replace("REJ", "");
						}else{
							tester=cell.getStringCellValue();
							assessment=null;
						}
					}else{
						tester=null;
						assessment=null;
					}
					ips.setString(78, tester);     //测试人
					ips.setString(79, assessment);   //评核
					
					setIPSValueReq(ips,rowm, 6, typePos[0],true);   //A面BGA要求
	         
					setIPSValueReq(ips,rowm,18, typePos[1],true);   //A面SMD要求
				           
					setIPSValueReq(ips,rowm,30, typePos[2],true);   //A面01005长要求
					         
					setIPSValueReq(ips,rowm,42, typePos[3],true);   //A面01005宽要求
	          
					setIPSValueReq(ips,rowm,54, typePos[4],true);   //A面0201长要求
	           
					setIPSValueReq(ips,rowm,66, typePos[5],true);   //A面0201宽要求
	        
					setIPSValueReq(ips,rowm,12, typePos[6],true);   //B面BGA要求
	    
					setIPSValueReq(ips,rowm,24, typePos[7],true);   //B面SMD要求
	    
					setIPSValueReq(ips,rowm,36, typePos[8],true);   //B面01005长要求
	    
					setIPSValueReq(ips,rowm,48, typePos[9],true);   //B面01005宽要求
	   
					setIPSValueReq(ips,rowm,60, typePos[10],true);  //B面0201长要求
	     
					setIPSValueReq(ips,rowm,72, typePos[11],true);  //B面0201宽要求
					
					ips.setString(80, remark);            //备注
					ips.setString(81, divCode);
					
					
					for(int j=0; j<5; j++){        		
						Row rown = sheet.getRow(i+j);
						if(rown!=null){						
						
							for(int F=0;F<typePos.length;F++){							
								switch (F) {
								case 0:
									   Interval=7;   //A面BGA实测值
									   break;
								case 1:
									   Interval=19;  //A面SMD实测值
									   break;
								case 2:
									   Interval=31;  //A面01005长实测值		
									   break;
								case 3:
									   Interval=43;   //A面01005宽实测值		
									   break;								   
								case 4:
									   Interval=55;   //A面0201长实测值
									   break;
								case 5:
									   Interval=67;   //A面0201宽实测值
									   break;	
								case 6:
									   Interval=13;   //B面BGA实测值
									   break;
								case 7:
									   Interval=25;   ////B面SMD实测值
									   break;	
								case 8:
									   Interval=37;   //B面01005长实测值
									   break;
								case 9:
									   Interval=49;   //B面01005宽实测值
									   break;
								case 10:
									   Interval=61;   //B面0201长实测值
									   break;
								case 11:
									   Interval=73;   //B面0201宽实测值
									   break;									   
								default:
									   Interval=7;
									break;
								}
								
								cell=rown.getCell(typePos[F]+3);	
								if(cell!=null&&cell.getCellType()==1){           
									if(LotNum.equals(cell.getStringCellValue())){
										setIPSValueReq(ips,rown,Interval+j, typePos[F],false);      
									}else{
										ips.setString(Interval+j, "0");	
									}
									
								}else{
									ips.setString(Interval+j, "0");
								}													
							}			
							
							/***********************************
							cell=rown.getCell(startPos+3);          //lot号
							if(cell!=null&&cell.getCellType()==1){ 
								if(LotNum.equals(cell.getStringCellValue())){		
									
									setIPSValueReq(ips,rown,7+j, typePos[0],false);      //A面BGA实测值	
									
									setIPSValueReq(ips,rown,19+j, typePos[1],false);     //A面SMD实测值

									setIPSValueReq(ips,rown,31+j, typePos[2],false);    //A面01005长实测值	
									
									setIPSValueReq(ips,rown,43+j, typePos[3],false);     //A面01005宽实测值	
	             
									setIPSValueReq(ips,rown,55+j, typePos[4],false);    //A面0201长实测值	
									
									setIPSValueReq(ips,rown,67+j, typePos[5],false);    //A面0201宽实测值		

									setIPSValueReq(ips,rown,13+j, typePos[6],false);      //B面BGA实测值		

									setIPSValueReq(ips,rown,25+j, typePos[7],false);     //B面SMD实测值

									setIPSValueReq(ips,rown,37+j, typePos[8],false);        //B面01005长实测值		

									setIPSValueReq(ips,rown,49+j, typePos[9],false);        //B面01005宽实测值

									setIPSValueReq(ips,rown,61+j, typePos[10],false);       //B面0201长实测值	
									
									setIPSValueReq(ips,rown,73+j, typePos[11],false);      //B面0201宽实测值	
									
								}else{
									for(int k=j; k<5;k++){
										ips.setString(7+k, "0");
										ips.setString(19+k, "0");
										ips.setString(31+k, "0");
										ips.setString(43+k, "0");
										ips.setString(55+k, "0");
										ips.setString(67+k, "0");
										ips.setString(13+k, "0");
										ips.setString(25+k, "0");
										ips.setString(37+k, "0");
										ips.setString(49+k, "0");
										ips.setString(61+k, "0");
										ips.setString(73+k, "0");
									}
									i=i-(5-j);
									break;
								}
							}else{
								for(int k=j; k<5;k++){
									ips.setString(7+k, "0");
									ips.setString(19+k, "0");
									ips.setString(31+k, "0");
									ips.setString(43+k, "0");
									ips.setString(55+k, "0");
									ips.setString(67+k, "0");
									ips.setString(13+k, "0");
									ips.setString(25+k, "0");
									ips.setString(37+k, "0");
									ips.setString(49+k, "0");
									ips.setString(61+k, "0");
									ips.setString(73+k, "0");
								}
								i=i-(5-j);
								break;
							}
						*/	
						}else{
							for(int k=j; k<5;k++){
								ips.setString(7+k, "0");
								ips.setString(19+k, "0");
								ips.setString(31+k, "0");
								ips.setString(43+k, "0");
								ips.setString(55+k, "0");
								ips.setString(67+k, "0");
								ips.setString(13+k, "0");
								ips.setString(25+k, "0");
								ips.setString(37+k, "0");
								ips.setString(49+k, "0");
								ips.setString(61+k, "0");
								ips.setString(73+k, "0");
							}
							break;
						}					
					}
					
					ips.addBatch();  
				    count = count + 1;
				}
			}catch (Exception e){
				e.printStackTrace();
				SendError(e.getMessage());
			}finally{
				return count;
			}	
		}	
		
		@SuppressWarnings("finally")
		public int setFQEPQA366AParam2(PreparedStatement ips, String monitorId,Sheet sheet,int fDatePOS,String divCode) throws Exception{
			int count=0;
			try{
				Row rowm = null;
				Cell cell = null;	
				String PN=null;       //型号
				String LotNum=null;   //lot号
				String remark=null;   //备注
				String LayerNum=null;
				Date dateTime=null;
				final int colInterval=7;
							 
				rowm = sheet.getRow(0);
				cell=rowm.getCell(1);
				
				if(cell!=null&&cell.getCellType()==1){           
					PN=cell.getStringCellValue();
					if(PN.contains("-")){
						LayerNum="L"+(PN.substring(PN.length()-4, PN.length()-3).equals("0")?PN.substring(PN.length()-3, PN.length()-2):PN.substring(PN.length()-4, PN.length()-2))       				
						          +"/"+(PN.substring(PN.length()-2,PN.length()-1).equals("0")?PN.substring(PN.length()-1):PN.substring(PN.length()-2));    //层次
						PN=PN.substring(0,PN.lastIndexOf('-'));
					}
				}	
				
				cell=rowm.getCell(4);			
				if(cell!=null&&cell.getCellType()==1){           
					remark=cell.getStringCellValue();        //备注固定填写在第一行第五列
				}	
				remark=LayerNum;           //层次填写在备注栏位
				
				int rowNum = sheet.getLastRowNum();
				for (int i = 3; i <= rowNum; i=i+5) {
					ips.setString(3, PN);
					
					rowm = sheet.getRow(i);				
					cell=rowm.getCell(4);				
					if(cell!=null&&cell.getCellType()==1){           //lot号
						LotNum=cell.getStringCellValue();
					}else{
						LotNum=null;
						break;
					}
					ips.setString(4, LotNum);
									
					cell=rowm.getCell(5);          //时间	
					if(cell!=null&&cell.getCellType()==0){           //日期和时间
						dateTime=cell.getDateCellValue();
						ips.setDate(1, new java.sql.Date(dateTime.getTime()));
						ips.setString(2, dateTime.getHours()+":"+dateTime.getMinutes());
					}else{
						ips.setString(1, null);
						ips.setString(2, null);
					}
					
					cell=rowm.getCell(41);       
					if(cell!=null&&cell.getCellType()==1){         
						ips.setString(41, cell.getStringCellValue());     //测试人
						ips.setString(42, cell.getStringCellValue().substring(0, 3));   //评核
					}else{
						ips.setString(41, null);     //测试人
						ips.setString(42, null);   //评核
					}

					for (int classifyNum=0; classifyNum<3; classifyNum++){
						cell=rowm.getCell(1+colInterval*classifyNum);					//A面单元内线宽要求、SMD直径要求、BGA直径要求
						setIPSValue(ips,cell,5+(colInterval-1)*classifyNum*2);
						cell=rowm.getCell(22+colInterval*classifyNum);					//B面单元内线宽要求、SMD直径要求、BGA直径要求
						setIPSValue(ips,cell,11+(colInterval-1)*classifyNum*2);	
						
						for(int j=0; j<5; j++){        			
							Row rown = sheet.getRow(i+j);
							if(rown!=null){
								cell=rown.getCell(4);	
								if(cell!=null&&cell.getCellType()==1){           
									if(LotNum.equals(cell.getStringCellValue())){
										cell=rown.getCell(2+colInterval*classifyNum);          //A面单元内线宽	、SMD直径、BGA直径
										setIPSValue1000(ips,cell,6+(colInterval-1)*classifyNum*2+j);
										cell=rown.getCell(23+colInterval*classifyNum);          //B面单元内线宽	、SMD直径、BGA直径	
										setIPSValue1000(ips,cell,12+(colInterval-1)*classifyNum*2+j);
									}else{
										for(int k=j; k<5;k++){
											ips.setString(6+(colInterval-1)*classifyNum*2+k, null);
											ips.setString(12+(colInterval-1)*classifyNum*2+k, null);
										}
										if(classifyNum==2) i=i-(5-j);
										break;
									}
								}else{
									for(int k=j; k<5;k++){
										ips.setString(6+(colInterval-1)*classifyNum*2+k, null);
										ips.setString(12+(colInterval-1)*classifyNum*2+k, null);
									}
									if(classifyNum==2) i=i-(5-j);
									break;
								}
							}else{
								for(int k=j; k<5;k++){
									ips.setString(6+(colInterval-1)*classifyNum*2+k, null);
									ips.setString(12+(colInterval-1)*classifyNum*2+k, null);								
								}
								if(classifyNum==2) i=i-(5-j);
								break;
							}
						}					
					}
					
					ips.setString(43, remark);            //备注
					ips.setString(44, divCode);
					ips.addBatch();
				    count = count + 1;
				}
			}catch (Exception e){
				e.printStackTrace();
				SendError(e.getMessage());
			}finally{
				return count;
			}	
		}
		@SuppressWarnings("finally")
		public int setDMELAM003AParam2(PreparedStatement ips, String monitorId,Sheet sheet,int fDatePOS,String divCode) throws Exception{
			int count=0;
			try{
				Row rowm = null;
				Cell cell = null;	
				String PN=null;       //型号
				String LotNum=null;   //lot号
				String remark=null;
				String layerNum=null;
				Date dateTime=null;
				
				rowm = sheet.getRow(0);
				
				cell=rowm.getCell(1);			
				if(cell!=null&&cell.getCellType()==1){           
					PN=cell.getStringCellValue();
					if(PN.contains("-")){
						layerNum="L"+(PN.substring(PN.length()-4, PN.length()-3).equals("0")?PN.substring(PN.length()-3, PN.length()-2):PN.substring(PN.length()-4, PN.length()-2))       				
								 +"/"+(PN.substring(PN.length()-2,PN.length()-1).equals("0")?PN.substring(PN.length()-1):PN.substring(PN.length()-2));    //层次
						PN=PN.substring(0,PN.lastIndexOf('-'));
					}
				}	
				cell=rowm.getCell(4);			
				if(cell!=null&&cell.getCellType()==1){           
					remark=cell.getStringCellValue();         //备注固定填写在第一行第五列
				}
				
				int rowNum = sheet.getLastRowNum();
				for (int i = 3; i <= rowNum; i=i+5) {
					ips.setString(3, PN);
					
					rowm = sheet.getRow(i);
					
					cell=rowm.getCell(4);				
					if(cell!=null&&cell.getCellType()==1){           //lot号
						LotNum=cell.getStringCellValue();
					}else{
						LotNum=null;
						break;
					}
					ips.setString(4, LotNum);
					
					ips.setString(5, layerNum);
					
					
					cell=rowm.getCell(5);
					if(cell!=null&&cell.getCellType()==0){           //日期和时间
						dateTime=cell.getDateCellValue();
						ips.setDate(1, new java.sql.Date(dateTime.getTime()));
						ips.setString(2, dateTime.getHours()+":"+dateTime.getMinutes());
					}else{
						ips.setString(1, null);
						ips.setString(2, null);
					}
					
					cell=rowm.getCell(82);      
					if(cell!=null&&cell.getCellType()==1){         
						ips.setString(78, cell.getStringCellValue());     //测试人
						ips.setString(79, cell.getStringCellValue().substring(0, 3));   //评核
					}else{
						ips.setString(78, null);     //测试人
						ips.setString(79, null);   //评核
					}
					
					cell=rowm.getCell(1);
					setIPSValue(ips,cell,6);           //A面BGA要求
					
					cell=rowm.getCell(8);
					setIPSValue(ips,cell,18);           //A面SMD要求
					
					cell=rowm.getCell(15);
					setIPSValue(ips,cell,30);           //A面01005长要求
					
					cell=rowm.getCell(22);
					setIPSValue(ips,cell,42);           //A面01005宽要求
					
					cell=rowm.getCell(29);
					setIPSValue(ips,cell,54);           //A面0201长要求
					
					cell=rowm.getCell(36);
					setIPSValue(ips,cell,66);           //A面0201宽要求
					
					cell=rowm.getCell(42);
					setIPSValue(ips,cell,12);           //B面BGA要求
					
					cell=rowm.getCell(49);
					setIPSValue(ips,cell,24);           //B面SMD要求
					
					cell=rowm.getCell(56);
					setIPSValue(ips,cell,36);           //B面01005长要求
					
					cell=rowm.getCell(63);
					setIPSValue(ips,cell,48);           //B面01005宽要求
					
					cell=rowm.getCell(70);
					setIPSValue(ips,cell,60);           //B面0201长要求
					
					cell=rowm.getCell(77);
					setIPSValue(ips,cell,72);           //B面0201宽要求
					
					ips.setString(80, remark);            //备注
					ips.setString(81, divCode);
					
					
					for(int j=0; j<5; j++){        		
						Row rown = sheet.getRow(i+j);
						if(rown!=null){
							cell=rown.getCell(4);          //lot号
							if(cell!=null&&cell.getCellType()==1){ 
								if(LotNum.equals(cell.getStringCellValue())){
									cell=rown.getCell(2);              //A面BGA实测值		
									setIPSValue(ips,cell,7+j);
									cell=rown.getCell(9);              //A面SMD实测值		
									setIPSValue(ips,cell,19+j);
									cell=rown.getCell(16);              //A面01005长实测值		
									setIPSValue(ips,cell,31+j);
									cell=rown.getCell(23);              //A面01005宽实测值		
									setIPSValue(ips,cell,43+j);
									cell=rown.getCell(30);              //A面0201长实测值		
									setIPSValue(ips,cell,55+j);
									cell=rown.getCell(37);              //A面0201宽实测值		
									setIPSValue(ips,cell,67+j);
									
									cell=rown.getCell(43);              //B面BGA实测值		
									setIPSValue(ips,cell,13+j);
									cell=rown.getCell(50);              //B面SMD实测值		
									setIPSValue(ips,cell,25+j);
									cell=rown.getCell(57);              //B面01005长实测值		
									setIPSValue(ips,cell,37+j);
									cell=rown.getCell(64);              //B面01005宽实测值		
									setIPSValue(ips,cell,49+j);
									cell=rown.getCell(71);              //B面0201长实测值		
									setIPSValue(ips,cell,61+j);
									cell=rown.getCell(78);              //B面0201宽实测值		
									setIPSValue(ips,cell,73+j);
									
								}else{
									for(int k=j; k<5;k++){
										ips.setString(7+k, null);
										ips.setString(19+k, null);
										ips.setString(31+k, null);
										ips.setString(43+k, null);
										ips.setString(55+k, null);
										ips.setString(67+k, null);
										ips.setString(13+k, null);
										ips.setString(25+k, null);
										ips.setString(37+k, null);
										ips.setString(49+k, null);
										ips.setString(61+k, null);
										ips.setString(73+k, null);
									}
									i=i-(5-j);
									break;
								}
							}else{
								for(int k=j; k<5;k++){
									ips.setString(7+k, null);
									ips.setString(19+k, null);
									ips.setString(31+k, null);
									ips.setString(43+k, null);
									ips.setString(55+k, null);
									ips.setString(67+k, null);
									ips.setString(13+k, null);
									ips.setString(25+k, null);
									ips.setString(37+k, null);
									ips.setString(49+k, null);
									ips.setString(61+k, null);
									ips.setString(73+k, null);
								}
								i=i-(5-j);
								break;
							}						
						}else{
							for(int k=j; k<5;k++){
								ips.setString(7+k, null);
								ips.setString(19+k, null);
								ips.setString(31+k, null);
								ips.setString(43+k, null);
								ips.setString(55+k, null);
								ips.setString(67+k, null);
								ips.setString(13+k, null);
								ips.setString(25+k, null);
								ips.setString(37+k, null);
								ips.setString(49+k, null);
								ips.setString(61+k, null);
								ips.setString(73+k, null);
							}
							break;
						}					
					}
					
					ips.addBatch();
				    count = count + 1;
				}
			}catch (Exception e){
				SendError(e.getMessage());
			}finally{
				return count;
			}	
		}
		

		@SuppressWarnings({ "finally", "deprecation" })
		public int ImpQCLDR01Data(PreparedStatement ips, String monitorId,
				Sheet sheet, int fDatePOS, String divCode) {
			int count=0;
			try{
				Row rowm = null;
				Cell cell = null;	
				Date dateTime=null;
				String DateStr;	 
			    SimpleDateFormat formatter = new SimpleDateFormat(Keys.DATE_TIME_FORMAT2);		  
			    DateFormat df = new SimpleDateFormat(Keys.DATE_TIME_FORMAT3,Locale.US);
			    
				rowm = sheet.getRow(0);				
				String PN = getCellValue(rowm.getCell(1));
				String Remark = getCellValue(rowm.getCell(9))+":"+getCellValue(rowm.getCell(10));
				
				rowm = sheet.getRow(2);	
				String Layer1 = getCellValue(rowm.getCell(1));
				String Layer2 = getCellValue(rowm.getCell(8));
			    
				String Lot="";
				int rowNum = sheet.getLastRowNum();
				int firstRow = 3;		
		
				if( firstRow >= rowNum){
					return -1;
				}
				for (int i =firstRow; i <= rowNum; i=i+5) {
					rowm = sheet.getRow(i);	
					if(!Lot.equals(getCellValue(rowm.getCell(4)))){
						Lot = getCellValue(rowm.getCell(4));
						cell=rowm.getCell(5);						
						if(cell ==null ){
							ips.setDate(1, new java.sql.Date(new Date().getTime()));
							ips.setString(2, "早班");					
						}else{ 
							DateStr = getCellValue(cell);					
							if(! DateStr.equals("")){ 
								if (DateStr.indexOf("CST") != -1){
									dateTime = df.parse(DateStr);
									DateStr = formatter.format(dateTime);
								}
								dateTime=formatter.parse(DateStr);
								ips.setDate(1, new java.sql.Date(dateTime.getTime()));
								ips.setString(2, dateTime.getHours()+":"+dateTime.getMinutes()); 
							}else{					
								ips.setDate(1, new java.sql.Date(new Date().getTime()));
								ips.setString(2, "早班");
							}	
						}
						
						ips.setString(3, PN);
						ips.setString(4, Lot);
						ips.setString(5, Layer1);
						
						for(int j=0; j<5; j++){         //实测值				
							Row rown = sheet.getRow(i+j);
							Cell subcell = rown.getCell(2);
							if(Lot.equals(getCellValue(rown.getCell(4)))){
							   setIPSValue(ips,subcell,6+j,3);			
							}else{
							   ips.setString(6+j, null);	
							}
						}				       								
						ips.setString(11, Remark);      //备注					
						ips.setString(12, divCode);
						ips.addBatch();
						
						if(!"".equals(Layer2)){
							cell=rowm.getCell(12);						
							if(cell ==null ){
								ips.setDate(1, new java.sql.Date(new Date().getTime()));
								ips.setString(2, "早班");					
							}else{ 
								DateStr = getCellValue(cell);					
								if(! DateStr.equals("")){        
									if (DateStr.indexOf("CST") != -1){
										dateTime = df.parse(DateStr);
										DateStr = formatter.format(dateTime);
									}							
									dateTime=formatter.parse(DateStr);
									ips.setDate(1, new java.sql.Date(dateTime.getTime()));
									ips.setString(2, dateTime.getHours()+":"+dateTime.getMinutes());
								}else{					
									ips.setDate(1, new java.sql.Date(new Date().getTime()));
									ips.setString(2, "早班");
								}	
							}
							
							ips.setString(3, PN);
							ips.setString(4, Lot);
							ips.setString(5, Layer2);
							
							for(int b=0; b<5; b++){         //实测值				
								Row rown = sheet.getRow(i+b);
								Cell subcell = rown.getCell(9);
								if(Lot.equals(getCellValue(rown.getCell(11)))){
								   setIPSValue(ips,subcell,6+b,3);			
								}else{
								   ips.setString(6+b, null);	
								}
							}				       
										
							ips.setString(11, Remark);      //备注
							
							ips.setString(12, divCode);
							ips.addBatch();
						}
						count = count + 1;				
				    }
				}
			}catch (Exception e){
				SendError(e.getMessage());			
			}finally{
				return count;
			}	
		}

		@SuppressWarnings("finally")
		public int ImpPQA374Data(PreparedStatement ips, String FileName,String monitorId,
				Sheet sheet, int fDatePOS, String divCode) {
			int count=0;
			try{
				String [] LotArray = FileName.split("-");
				String PN,Lot,Level;
				if(LotArray.length>=4){
					PN = LotArray[0];
					if(PN.indexOf("_")>=1){
						PN =PN.substring(PN.indexOf("_")+1);
					}
					Lot = LotArray[1] +"-"+LotArray[2];
					Level = LotArray[3].substring(0, 4);
					Lot = Lot+"-"+ Level;
					if("0000".equals(Level)){
						Level = "L1/"+PN.substring(0, 2);
					}else{
						Level = "L"+ Integer.parseInt(Level.substring(0,2))+"/"+Integer.parseInt(Level.substring(2));
					}
					
					
					int rowNum = sheet.getLastRowNum();
					int firstRow = 10;	
					if( firstRow >= rowNum){
						return -1;
					}
					Row row = null;		
					String cellValue,Line="#";
					
					Date dateTime=null;
					String DateStr,CheckUser="",Remark="";	
				    SimpleDateFormat formatter = new SimpleDateFormat(Keys.DATE_TIME_FORMAT2);		  
				    DateFormat df = new SimpleDateFormat(Keys.DATE_TIME_FORMAT3,Locale.US);
				    
					for (int i =firstRow; i <= rowNum; i++) {
						row = sheet.getRow(i);	
						cellValue = getCellValue(row.getCell(0));
						if(cellValue.contains("日期")){
							DateStr = getCellValue(row.getCell(5));					
							if(! DateStr.equals("")){ 
								if (DateStr.indexOf("CST") != -1){
									dateTime = df.parse(DateStr);
									DateStr = formatter.format(dateTime);
								}
								//"2014-3-26:8:24:57"
								DateStr = DateStr.replaceAll("-", "/");
								DateStr = DateStr.replaceFirst(":", " ");
								dateTime=formatter.parse(DateStr);						
							}
						}else if(cellValue.contains("测量部门")){
							Remark  =cellValue +":"+ getCellValue(row.getCell(5));
						}else if(cellValue.contains("测量模式")){
							Remark = Remark +"  "+ cellValue +":"+ getCellValue(row.getCell(5));
						}else if(cellValue.contains("测量人")){
							CheckUser = getCellValue(row.getCell(5));
						}else if(cellValue.contains("1") || cellValue.contains("2")){
							if(dateTime !=null){	
								ips.setDate(1, new java.sql.Date(dateTime.getTime()));
								ips.setString(2, dateTime.getHours()+":"+dateTime.getMinutes());
							}else{					
								ips.setDate(1, new java.sql.Date(new Date().getTime()));
								ips.setString(2, "早班");
							}	
							
							if(cellValue.contains("1")){
								Line =getCellValue(row.getCell(7));
							}
							
							if("#".equals(Line)){
							   ips.setString(3, "/");
							}else{
							   ips.setString(3, Line);
							}
							ips.setString(4, PN);
							ips.setString(5, Lot);
							ips.setString(6, Level);
							ips.setString(7, "/");
							ips.setString(8, "/");
							ips.setString(9, getRoundNum(getCellValue(row.getCell(1)),2));
							ips.setString(10, getRoundNum(getCellValue(row.getCell(2)),2));
							ips.setString(11, getRoundNum(getCellValue(row.getCell(3)),2));
							ips.setString(12, getRoundNum(getCellValue(row.getCell(4)),2));
							ips.setString(13, getRoundNum(getCellValue(row.getCell(5)),2));
							
							ips.setString(14, "/");
							ips.setString(15, "Acc");
							ips.setString(16, CheckUser);
							ips.setString(17, Remark);
							ips.setString(18, divCode);
							ips.addBatch();
							count = count + 1;
						}else{
							continue;
						}
					}
				}
				
			}catch (Exception e){
				SendError(e.getMessage());			
			}finally{
				return count;
			}
		}	
       
		//410 new 	
		 private int setQEPQA410AParam(PreparedStatement ips, String monitorId, Sheet sheet, int fDatePOS, String divCode)
				    throws Exception{
			   int count = 0;
			    Connection conn = null;
			    PreparedStatement psMast = null;
			    PreparedStatement detlPs = null;
			    PreparedStatement psMastPkey = null;
			    PreparedStatement psMastSeriesPkey = null;
			    PreparedStatement psSeriesMast = null;
			    PreparedStatement psSeriesDetl = null;

			    ResultSet rs = null;
			    ResultSet rsSeries = null;
			    try {
			      Row rowm = null;
			      Cell cell = null;
			      String machNO = monitorId;

			      conn = DBOptionUtil.getConnection();

			      String pn = null;
			      String wo = null;
			      rowm = sheet.getRow(0);

			      cell = rowm.getCell(1);
			      if ((cell != null) && (cell.getCellType() == 1)) {
			        pn = cell.getStringCellValue();
			      }

			      rowm = sheet.getRow(2);
			      List list = new ArrayList();
			      for (int i = 0; i < 256; i += 7) {
			        cell = rowm.getCell(i + 1);
			        if ((cell != null) && (cell.getCellType() == 1)) {
			          list.add(Integer.valueOf(cell.getColumnIndex()));
			        }
			      }

			      if ((list != null) && (!list.isEmpty()) && (list.size() > 0)) {
			        int rowNum = sheet.getLastRowNum();
			        int first_m = -1;
			        for (int j = 0; j < list.size(); j++) {
			          if (j == 0) {
			            for (int i = 2; i <= rowNum; i++) {
			              rowm = sheet.getRow(i);
			              int n = 0;
			              for (int k = Integer.parseInt(((Integer)list.get(j)).toString()); k <= Integer.parseInt(((Integer)list.get(j)).toString()) + 4; k++) {
			                cell = rowm.getCell(k);
			                if ((n == 0) && 
			                  (cell != null) && (cell.getCellType() == 1) && 
			                  (cell.getStringCellValue() != null)) {
			                  first_m++;
			                }

			                n++;
			              }
			            }

			          }

			        }

			        BigDecimal mastPkey = null;
			        StringBuffer detlSql = new StringBuffer();

			        String sqlMast = "INSERT INTO MTG_IPQC_TRAC_MAST(PKEY,PN,WO,LINE_NO,LAYER,STATUS,TDATE,CURRENT_VERSION,DIV_CODE) VALUES(:PKEY,:PN,:WO,:LINE_NO,:LAYER,:STATUS,:tDate,0,:DIV_CODE)";

			        detlSql.append("INSERT INTO MTG_IPQC_TRAC_DETL(PKEY,MTG_IPQC_TRAC_MAST_PTR,STEP,TTYPE,LINE_WIDTH,DESD_NO,UNIT_1,UNIT_2,UNIT_3,UNIT_4,UNIT_5,IPQC_AUTO_JUDGE,DESC_1,DESC_2,DESC_3,DESC_4,EDIT_BY,IPQC_APPE_STATUS,IPQC_REMARK,SPC_LINE_WIDTH,SPC_STATUS,TDATE,LAST_EDIT_BY,EDIT_DATE,CURRENT_VERSION,DIV_CODE) ");

			        detlSql.append(" VALUES(MTG_IPQC_TRAC_DETL_SEQ.NEXTVAL,:MTG_IPQC_TRAC_MAST_PTR,:STEP,:TTYPE,:LINE_WIDTH,:DESD_NO,:UNIT_1,:UNIT_2,:UNIT_3,:UNIT_4,:UNIT_5,:IPQC_AUTO_JUDGE,:DESC_1,:DESC_2,:DESC_3,:DESC_4,:EDIT_BY,:IPQC_APPE_STATUS,:IPQC_REMARK,:SPC_LINE_WIDTH,:SPC_STATUS,:tdate,:LAST_EDIT_BY,sysdate,0,:DIV_CODE)");

			        String seriesMastSql = "INSERT INTO SPC_COPPERTHICK_MAST(PKEY,MAIN_CATEGORY,SUB_ITEM,TRAN_STATUS,SPC_DATE,CURR_CLASS,MACHINE_NO,SERIES_NO,PART_NUMBER,LOT_CARD,CURR_LAYER, SAMPLE_VOLUME,USL,TARGET,LSL,DIRECTION,REMARK,SPC_TIME) VALUES(:mastPkey,'01-IDF','Anti pad尺寸(AP3)',1,:SPC_DATE,:CURR_CLASS,:MACHINE_NO,:SERIES_NO,:PART_NUMBER,:LOT_CARD,:CURR_LAYER,0,null,null,null,null,null,:SPC_TIME)";

			        String seriesDetlSql = "INSERT INTO SPC_COPPERTHICK_DETL(THICK_PTR,INSPECT_ITEMS,STEP,MEASURED_VALUE)VALUES(:mastPkey,'Anti pad尺寸',:step,:detlValue)";

			        psMast = conn.prepareStatement(sqlMast);
			        detlPs = conn.prepareStatement(detlSql.toString());
			        psSeriesMast = conn.prepareStatement(seriesMastSql);
			        psSeriesDetl = conn.prepareStatement(seriesDetlSql);

			        String workOrderNumber = "";
			        String layer = null;
			        String floor = null;
			        java.util.Date globalDateTime = null;
			        Boolean executeFlag = Boolean.valueOf(false);
			        for (int j = 0; j < list.size(); j++) {
			          String Linewidth = null;
			          String ipqc_check = "ACC";
			          String desdNo = null;
			          String unit1 = null;
			          String unit2 = null;
			          String unit3 = null;
			          String unit4 = null;
			          String unit5 = null;
			          String testBy = null;
			          String audit = null;
			          String remark = null;
			          java.util.Date dateTime = null;
			          String ttype = null;
			          if (j == 0) {
			            int n = 0;
			            int unit_row = 0;
			            for (int i = 2; i <= rowNum; i++) {
			              rowm = sheet.getRow(i);
			              int column_index = 0;
			              for (int k = Integer.parseInt(((Integer)list.get(j)).toString()); k <= Integer.parseInt(((Integer)list.get(j)).toString()) + 4; k++) {
			                cell = rowm.getCell(k);
			                column_index++;
			                if ((n == 0) && 
			                  (cell != null) && (cell.getCellType() == 1) && 
			                  (cell.getStringCellValue() != null) && (!"".equals(cell.getStringCellValue()))) {
			                  if (cell.getStringCellValue().indexOf("-") != -1) {
			                    layer = cell.getStringCellValue().split("-")[0];
			                    ttype = cell.getStringCellValue().substring(cell.getStringCellValue().split("-")[0].length() + 1, cell.getStringCellValue().length());
			                  } else {
			                    layer = cell.getStringCellValue().indexOf("L") != -1 ? cell.getStringCellValue() : "";
			                    ttype = cell.getStringCellValue().indexOf("L") == -1 ? cell.getStringCellValue() : "";
			                  }

			                }

			                if ((first_m == 2) || (first_m == 6)) {
			                  if ((column_index == 1) && (cell != null) && (cell.getCellType() == 1)) {
			                    Linewidth = cell.getStringCellValue();
			                  }
			                  if ((column_index == 2) && (cell != null) && (cell.getCellType() == 0)) {
			                    unit_row++;
			                    if (unit_row == 1)
			                      desdNo = String.valueOf(cell.getNumericCellValue());
			                    else if (unit_row == 2)
			                      unit1 = String.valueOf(cell.getNumericCellValue());
			                    else if (unit_row == 3)
			                      unit2 = String.valueOf(cell.getNumericCellValue());
			                    else if (unit_row == 4)
			                      unit3 = String.valueOf(cell.getNumericCellValue());
			                    else if (unit_row == 5)
			                      unit4 = String.valueOf(cell.getNumericCellValue());
			                    else if (unit_row == 6) {
			                      unit5 = String.valueOf(cell.getNumericCellValue());
			                    }
			                  }
			                  if ((column_index == 3) && (cell != null) && (cell.getCellType() == 0) && 
			                    (cell.getNumericCellValue() != 0.0D)) {
			                    ipqc_check = "REJ";
			                  }

			                  if ((column_index == 4) && (cell != null) && (cell.getCellType() == 1)) {
			                    wo = cell.getStringCellValue();
			                  }

			                  if ((column_index == 5) && (cell != null) && (cell.getCellType() == 0))
			                    dateTime = cell.getDateCellValue();
			                }
			                else {
			                  if ((column_index == 1) && (cell != null) && (cell.getCellType() == 1)) {
			                    Linewidth = cell.getStringCellValue();
			                  }
			                  if ((column_index == 2) && (cell != null) && (cell.getCellType() == 0)) {
			                    unit_row++;
			                    if (unit_row == 1)
			                      unit1 = String.valueOf(cell.getNumericCellValue());
			                    else if (unit_row == 2)
			                      unit2 = String.valueOf(cell.getNumericCellValue());
			                    else if (unit_row == 3)
			                      unit3 = String.valueOf(cell.getNumericCellValue());
			                    else if (unit_row == 4)
			                      unit4 = String.valueOf(cell.getNumericCellValue());
			                    else if (unit_row == 5) {
			                      unit5 = String.valueOf(cell.getNumericCellValue());
			                    }
			                  }
			                  if ((column_index == 3) && (cell != null) && (cell.getCellType() == 0) && 
			                    (cell.getNumericCellValue() != 0.0D)) {
			                    ipqc_check = "REJ";
			                  }

			                  if ((column_index == 4) && (cell != null) && (cell.getCellType() == 1)) {
			                    wo = cell.getStringCellValue();
			                  }

			                  if ((column_index == 5) && (cell != null) && (cell.getCellType() == 0)) {
			                    dateTime = cell.getDateCellValue();
			                  }
			                }
			              }
			              n++;
			            }

			            if (desdNo != null) {
			              if (desdNo.endsWith(".0"))
			                desdNo = desdNo.substring(0, desdNo.length() - 2);
			              else if (desdNo.contains(".")) {
			                desdNo = getRoundNum(desdNo);
			              }
			            }
			            if (unit1 != null) {
			              if (unit1.endsWith(".0"))
			                unit1 = unit1.substring(0, unit1.length() - 2);
			              else if (unit1.contains(".")) {
			                unit1 = getRoundNum(unit1);
			              }
			            }
			            if (unit2 != null) {
			              if (unit2.endsWith(".0"))
			                unit2 = unit2.substring(0, unit2.length() - 2);
			              else if (unit2.contains(".")) {
			                unit2 = getRoundNum(unit2);
			              }
			            }
			            if (unit3 != null) {
			              if (unit3.endsWith(".0"))
			                unit3 = unit3.substring(0, unit3.length() - 2);
			              else if (unit3.contains(".")) {
			                unit3 = getRoundNum(unit3);
			              }
			            }
			            if (unit4 != null) {
			              if (unit4.endsWith(".0"))
			                unit4 = unit4.substring(0, unit4.length() - 2);
			              else if (unit4.contains(".")) {
			                unit4 = getRoundNum(unit4);
			              }
			            }
			            if (unit5 != null) {
			              if (unit5.endsWith(".0"))
			                unit5 = unit5.substring(0, unit5.length() - 2);
			              else if (unit5.contains(".")) {
			                unit5 = getRoundNum(unit5);
			              }
			            }
			            workOrderNumber = wo.split("-").length > 4 ? wo.substring(0, wo.length() - wo.split("-")[4].length() - 1) : wo;
			            floor = workOrderNumber.substring(workOrderNumber.length() - 4, workOrderNumber.length()).substring(0, 2);
			            testBy = wo.split("-").length > 5 ? wo.substring(0, wo.length() - wo.split("-")[5].length() - 1) : null;
			            globalDateTime = dateTime;
			            if (dateTime == null)
			            {
			              continue;
			            }

			            ips.setDate(1, new java.sql.Date(dateTime.getTime()));

			            ips.setString(2, new SimpleDateFormat("HH:mm").format(Long.valueOf(dateTime.getTime())));
			            ips.setString(3, pn.split("-").length > 1 ? pn.split("-")[0] : pn);
			            ips.setString(4, wo.split("-").length > 4 ? wo.substring(0, wo.length() - wo.split("-")[4].length() - 1) : wo);
			            ips.setString(5, layer);
			            ips.setString(6, Linewidth);
			            ips.setString(7, unit1);
			            ips.setString(8, unit2);
			            ips.setString(9, unit3);
			            ips.setString(10, unit4);
			            ips.setString(11, unit5);
			            ips.setString(12, wo.split("-").length > 4 ? wo.split("-")[4] : "");
			            ips.setString(13, ipqc_check);
			            ips.setString(14, null);
			            ips.setString(15, null);

			            if ((unit1 != null) && (unit2 != null) && (unit3 != null) && (unit4 != null) && (unit5 != null)) {
			              List spclist = searchSpcControlInfo(pn.split("-").length > 1 ? pn.split("-")[0] : pn, layer, floor, divCode);
			              BigDecimal CONTROL_MIN = null;
			              BigDecimal CONTROL_MAX = null;
			              BigDecimal LIMIT_MAX = null;
			              if ((spclist != null) && (spclist.size() > 0)) {
			                CONTROL_MIN = new BigDecimal(((Object[])spclist.get(0))[0].toString());
			                CONTROL_MAX = new BigDecimal(((Object[])spclist.get(0))[1].toString());
			                LIMIT_MAX = new BigDecimal(((Object[])spclist.get(0))[2].toString());
			                String checkFlag = "Y";
			                BigDecimal mean = new BigDecimal(unit1).add(new BigDecimal(unit2)).add(new BigDecimal(unit3)).add(new BigDecimal(unit4)).add(new BigDecimal(unit5)).divide(new BigDecimal(5));
			                if ((CONTROL_MIN.compareTo(mean) > 0) || (CONTROL_MAX.compareTo(mean) < 0)) {
			                  checkFlag = "N";
			                }
			                BigDecimal[] array = { new BigDecimal(unit1), new BigDecimal(unit2), new BigDecimal(unit3), new BigDecimal(unit4), new BigDecimal(unit5) };
			                BigDecimal limitValue = getLimitValue(array);
			                if (limitValue.compareTo(LIMIT_MAX) > 0) {
			                  checkFlag = "N";
			                }
			                ips.setString(16, CONTROL_MIN + "--" + CONTROL_MAX);
			                ips.setString(17, checkFlag);
			              } else {
			                ips.setString(16, "--");
			                ips.setString(17, "--");
			              }
			              if ("PAD".equals(ttype)) {
			                String seriesNo = getPNSeries(pn.indexOf("-") > -1 ? pn.substring(0, pn.indexOf("-") - 3) : pn.substring(0, pn.length() - 3), "GME");
			                if (seriesNo != null) {
			                  executeFlag = Boolean.valueOf(true);
			                  psMastSeriesPkey = conn.prepareStatement("SELECT SPC_COPPERTHICK_MAST_SEQ.NEXTVAL FROM DUAL ");
			                  rsSeries = psMastSeriesPkey.executeQuery();
			                  rsSeries.next();
			                  BigDecimal mastSeriesPkey = rsSeries.getBigDecimal("NEXTVAL");
			                  psSeriesMast.setBigDecimal(1, mastSeriesPkey);
			                  psSeriesMast.setTimestamp(2, new Timestamp(dateTime.getTime()));
			                  psSeriesMast.setString(3, new SimpleDateFormat("HH:mm").format(Long.valueOf(dateTime.getTime())));
			                  psSeriesMast.setString(4, workOrderNumber.substring(workOrderNumber.length() - 4, workOrderNumber.length()));
			                  psSeriesMast.setString(5, seriesNo);
			                  psSeriesMast.setString(6, pn.split("-").length > 1 ? pn.split("-")[0] : pn);
			                  psSeriesMast.setString(7, workOrderNumber.substring(0, workOrderNumber.length() - 5));
			                  psSeriesMast.setString(8, layer);
			                  psSeriesMast.setTimestamp(9, new Timestamp(dateTime.getTime()));
			                  psSeriesMast.addBatch();

			                  psSeriesDetl.setBigDecimal(1, mastSeriesPkey);
			                  for (int ii = 1; ii <= 5; ii++)
			                    if (ii == 1) {
			                      psSeriesDetl.setBigDecimal(2, new BigDecimal(ii));
			                      psSeriesDetl.setBigDecimal(3, new BigDecimal(unit1));
			                      psSeriesDetl.addBatch();
			                    } else if (ii == 2) {
			                      psSeriesDetl.setBigDecimal(2, new BigDecimal(ii));
			                      psSeriesDetl.setBigDecimal(3, new BigDecimal(unit2));
			                      psSeriesDetl.addBatch();
			                    } else if (ii == 3) {
			                      psSeriesDetl.setBigDecimal(2, new BigDecimal(ii));
			                      psSeriesDetl.setBigDecimal(3, new BigDecimal(unit3));
			                      psSeriesDetl.addBatch();
			                    } else if (ii == 4) {
			                      psSeriesDetl.setBigDecimal(2, new BigDecimal(ii));
			                      psSeriesDetl.setBigDecimal(3, new BigDecimal(unit4));
			                      psSeriesDetl.addBatch();
			                    } else if (ii == 5) {
			                      psSeriesDetl.setBigDecimal(2, new BigDecimal(ii));
			                      psSeriesDetl.setBigDecimal(3, new BigDecimal(unit5));
			                      psSeriesDetl.addBatch();
			                    }
			                }
			              }
			            }
			            else {
			              ips.setString(16, "--");
			              ips.setString(17, "--");
			            }

			            ips.setString(18, divCode);

			            psMastPkey = conn.prepareStatement("SELECT MTG_IPQC_TRAC_MAST_SEQ.NEXTVAL FROM DUAL ");
			            rs = psMastPkey.executeQuery();
			            rs.next();
			            mastPkey = rs.getBigDecimal("NEXTVAL");

			            psMast.setBigDecimal(1, mastPkey);
			            psMast.setString(2, pn.split("-").length > 1 ? pn.split("-")[0] : pn);
			            psMast.setString(3, workOrderNumber.substring(0, workOrderNumber.length() - 5));
			            psMast.setString(4, workOrderNumber.substring(workOrderNumber.length() - 4, workOrderNumber.length()));
			            psMast.setString(5, layer);
			            psMast.setString(6, null);
			            psMast.setTimestamp(7, new Timestamp(dateTime.getTime()));
			            psMast.setString(8, divCode);

			            psMast.addBatch();

			            detlPs.setBigDecimal(1, mastPkey);
			            detlPs.setLong(2, count + 1);
			            detlPs.setString(3, ttype);
			            detlPs.setString(4, Linewidth);
			            detlPs.setString(5, desdNo);
			            detlPs.setString(6, unit1);
			            detlPs.setString(7, unit2);
			            detlPs.setString(8, unit3);
			            detlPs.setString(9, unit4);
			            detlPs.setString(10, unit5);
			            detlPs.setString(11, "ACC".equals(ipqc_check) ? "Y" : "N");
			            detlPs.setString(12, null);
			            detlPs.setString(13, null);
			            detlPs.setString(14, null);
			            detlPs.setString(15, null);
			            detlPs.setString(16, wo.split("-").length > 4 ? wo.split("-")[4] : "");
			            detlPs.setString(17, null);
			            detlPs.setString(18, null);

			            if ((unit1 != null) && (unit2 != null) && (unit3 != null) && (unit4 != null) && (unit5 != null)) {
			              List spclist = searchSpcControlInfo(pn.split("-").length > 1 ? pn.split("-")[0] : pn, layer, floor, divCode);
			              BigDecimal CONTROL_MIN = null;
			              BigDecimal CONTROL_MAX = null;
			              BigDecimal LIMIT_MAX = null;
			              if ((spclist != null) && (spclist.size() > 0)) {
			                CONTROL_MIN = new BigDecimal(((Object[])spclist.get(0))[0].toString());
			                CONTROL_MAX = new BigDecimal(((Object[])spclist.get(0))[1].toString());
			                LIMIT_MAX = new BigDecimal(((Object[])spclist.get(0))[2].toString());
			                String checkFlag = "Y";
			                BigDecimal mean = new BigDecimal(unit1).add(new BigDecimal(unit2)).add(new BigDecimal(unit3)).add(new BigDecimal(unit4)).add(new BigDecimal(unit5)).divide(new BigDecimal(5));
			                if ((CONTROL_MIN.compareTo(mean) > 0) || (CONTROL_MAX.compareTo(mean) < 0)) {
			                  checkFlag = "N";
			                }
			                BigDecimal[] array = { new BigDecimal(unit1), new BigDecimal(unit2), new BigDecimal(unit3), new BigDecimal(unit4), new BigDecimal(unit5) };
			                BigDecimal limitValue = getLimitValue(array);
			                if (limitValue.compareTo(LIMIT_MAX) > 0) {
			                  checkFlag = "N";
			                }
			                detlPs.setString(19, CONTROL_MIN + "--" + CONTROL_MAX);
			                detlPs.setString(20, checkFlag);
			              } else {
			                detlPs.setString(19, "--");
			                detlPs.setString(20, "--");
			              }
			            } else {
			              detlPs.setString(19, "--");
			              detlPs.setString(20, "--");
			            }
			            detlPs.setTimestamp(21, new Timestamp(dateTime.getTime()));
			            detlPs.setBigDecimal(22, null);
			            detlPs.setString(23, divCode);

			            detlPs.addBatch();
			            count++;
			          } else {
			            int n = 0;
			            int unit_row = 0;
			            for (int i = 2; i <= rowNum; i++) {
			              rowm = sheet.getRow(i);
			              int column_index = 0;
			              for (int k = Integer.parseInt(((Integer)list.get(j)).toString()); k <= Integer.parseInt(((Integer)list.get(j)).toString()) + 4; k++) {
			                cell = rowm.getCell(k);
			                column_index++;
			                if ((n == 0) && 
			                  (cell != null) && (cell.getCellType() == 1) && 
			                  (cell.getStringCellValue() != null) && (!"".equals(cell.getStringCellValue()))) {
			                  if (cell.getStringCellValue().indexOf("-") != -1) {
			                    layer = cell.getStringCellValue().split("-")[0];
			                    ttype = cell.getStringCellValue().substring(cell.getStringCellValue().split("-")[0].length() + 1, cell.getStringCellValue().length());
			                  } else {
			                    ttype = cell.getStringCellValue();
			                  }

			                }

			                if ((column_index == 1) && (cell != null) && (cell.getCellType() == 1)) {
			                  Linewidth = cell.getStringCellValue();
			                }
			                if ((column_index == 2) && (cell != null) && (cell.getCellType() == 0)) {
			                  unit_row++;

			                  if (unit_row == 1)
			                    unit1 = String.valueOf(cell.getNumericCellValue());
			                  else if (unit_row == 2)
			                    unit2 = String.valueOf(cell.getNumericCellValue());
			                  else if (unit_row == 3)
			                    unit3 = String.valueOf(cell.getNumericCellValue());
			                  else if (unit_row == 4)
			                    unit4 = String.valueOf(cell.getNumericCellValue());
			                  else if (unit_row == 5) {
			                    unit5 = String.valueOf(cell.getNumericCellValue());
			                  }
			                }
			                if ((column_index == 3) && (cell != null) && (cell.getCellType() == 0) && 
			                  (cell.getNumericCellValue() != 0.0D)) {
			                  ipqc_check = "REJ";
			                }

			                if ((column_index == 4) && (cell != null) && (cell.getCellType() == 1)) {
			                  wo = cell.getStringCellValue();
			                }
			                if ((column_index == 5) && (cell != null) && (cell.getCellType() == 0)) {
			                  dateTime = cell.getDateCellValue();
			                }
			              }
			              n++;
			            }
			            if (desdNo != null) {
			              if (desdNo.endsWith(".0"))
			                desdNo = desdNo.substring(0, desdNo.length() - 2);
			              else if (desdNo.contains(".")) {
			                desdNo = getRoundNum(desdNo);
			              }
			            }
			            if (unit1 != null) {
			              if (unit1.endsWith(".0"))
			                unit1 = unit1.substring(0, unit1.length() - 2);
			              else if (unit1.contains(".")) {
			                unit1 = getRoundNum(unit1);
			              }
			            }
			            if (unit2 != null) {
			              if (unit2.endsWith(".0"))
			                unit2 = unit2.substring(0, unit2.length() - 2);
			              else if (unit2.contains(".")) {
			                unit2 = getRoundNum(unit2);
			              }
			            }
			            if (unit3 != null) {
			              if (unit3.endsWith(".0"))
			                unit3 = unit3.substring(0, unit3.length() - 2);
			              else if (unit3.contains(".")) {
			                unit3 = getRoundNum(unit3);
			              }
			            }
			            if (unit4 != null) {
			              if (unit4.endsWith(".0"))
			                unit4 = unit4.substring(0, unit4.length() - 2);
			              else if (unit4.contains(".")) {
			                unit4 = getRoundNum(unit4);
			              }
			            }
			            if (unit5 != null) {
			              if (unit5.endsWith(".0"))
			                unit5 = unit5.substring(0, unit5.length() - 2);
			              else if (unit5.contains(".")) {
			                unit5 = getRoundNum(unit5);
			              }
			            }
			            if (dateTime == null)
			            {
			              continue;
			            }

			            detlPs.setBigDecimal(1, mastPkey);
			            detlPs.setLong(2, count + 1);
			            detlPs.setString(3, ttype);
			            detlPs.setString(4, Linewidth);
			            detlPs.setString(5, desdNo);
			            detlPs.setString(6, unit1);
			            detlPs.setString(7, unit2);
			            detlPs.setString(8, unit3);
			            detlPs.setString(9, unit4);
			            detlPs.setString(10, unit5);
			            detlPs.setString(11, "ACC".equals(ipqc_check) ? "Y" : "N");
			            detlPs.setString(12, null);
			            detlPs.setString(13, null);
			            detlPs.setString(14, null);
			            detlPs.setString(15, null);
			            detlPs.setString(16, wo.split("-").length > 4 ? wo.split("-")[4] : "");
			            detlPs.setString(17, null);
			            detlPs.setString(18, null);
			            if ((unit1 != null) && (unit2 != null) && (unit3 != null) && (unit4 != null) && (unit5 != null)) {
			              List spclist = searchSpcControlInfo(pn.split("-").length > 1 ? pn.split("-")[0] : pn, layer, floor, divCode);
			              BigDecimal CONTROL_MIN = null;
			              BigDecimal CONTROL_MAX = null;
			              BigDecimal LIMIT_MAX = null;
			              if ((spclist != null) && (spclist.size() > 0)) {
			                CONTROL_MIN = new BigDecimal(((Object[])spclist.get(0))[0].toString());
			                CONTROL_MAX = new BigDecimal(((Object[])spclist.get(0))[1].toString());
			                LIMIT_MAX = new BigDecimal(((Object[])spclist.get(0))[2].toString());
			                String checkFlag = "Y";
			                BigDecimal mean = new BigDecimal(unit1).add(new BigDecimal(unit2)).add(new BigDecimal(unit3)).add(new BigDecimal(unit4)).add(new BigDecimal(unit5)).divide(new BigDecimal(5));
			                if ((CONTROL_MIN.compareTo(mean) > 0) || (CONTROL_MAX.compareTo(mean) < 0)) {
			                  checkFlag = "N";
			                }
			                BigDecimal[] array = { new BigDecimal(unit1), new BigDecimal(unit2), new BigDecimal(unit3), new BigDecimal(unit4), new BigDecimal(unit5) };
			                BigDecimal limitValue = getLimitValue(array);
			                if (limitValue.compareTo(LIMIT_MAX) > 0) {
			                  checkFlag = "N";
			                }
			                detlPs.setString(19, CONTROL_MIN + "--" + CONTROL_MAX);
			                detlPs.setString(20, checkFlag);
			              } else {
			                detlPs.setString(19, "--");
			                detlPs.setString(20, "--");
			              }
			              if ("PAD".equals(ttype)) {
			                String seriesNo = getPNSeries(pn.indexOf("-") > -1 ? pn.substring(0, pn.indexOf("-") - 3) : pn.substring(0, pn.length() - 3), "GME");
			                if (seriesNo != null) {
			                  executeFlag = Boolean.valueOf(true);
			                  psMastSeriesPkey = conn.prepareStatement("SELECT SPC_COPPERTHICK_MAST_SEQ.NEXTVAL FROM DUAL ");
			                  rsSeries = psMastSeriesPkey.executeQuery();
			                  rsSeries.next();
			                  BigDecimal mastSeriesPkey = rsSeries.getBigDecimal("NEXTVAL");
			                  psSeriesMast.setBigDecimal(1, mastSeriesPkey);
			                  psSeriesMast.setTimestamp(2, new Timestamp(dateTime.getTime()));
			                  psSeriesMast.setString(3, new SimpleDateFormat("HH:mm").format(Long.valueOf(dateTime.getTime())));
			                  psSeriesMast.setString(4, workOrderNumber.substring(workOrderNumber.length() - 4, workOrderNumber.length()));
			                  psSeriesMast.setString(5, seriesNo);
			                  psSeriesMast.setString(6, pn.split("-").length > 1 ? pn.split("-")[0] : pn);
			                  psSeriesMast.setString(7, workOrderNumber.substring(0, workOrderNumber.length() - 5));
			                  psSeriesMast.setString(8, layer);
			                  psSeriesMast.setTimestamp(9, new Timestamp(dateTime.getTime()));
			                  psSeriesMast.addBatch();

			                  psSeriesDetl.setBigDecimal(1, mastSeriesPkey);
			                  for (int ii = 1; ii <= 5; ii++)
			                    if (ii == 1) {
			                      psSeriesDetl.setBigDecimal(2, new BigDecimal(ii));
			                      psSeriesDetl.setBigDecimal(3, new BigDecimal(unit1));
			                      psSeriesDetl.addBatch();
			                    } else if (ii == 2) {
			                      psSeriesDetl.setBigDecimal(2, new BigDecimal(ii));
			                      psSeriesDetl.setBigDecimal(3, new BigDecimal(unit2));
			                      psSeriesDetl.addBatch();
			                    } else if (ii == 3) {
			                      psSeriesDetl.setBigDecimal(2, new BigDecimal(ii));
			                      psSeriesDetl.setBigDecimal(3, new BigDecimal(unit3));
			                      psSeriesDetl.addBatch();
			                    } else if (ii == 4) {
			                      psSeriesDetl.setBigDecimal(2, new BigDecimal(ii));
			                      psSeriesDetl.setBigDecimal(3, new BigDecimal(unit4));
			                      psSeriesDetl.addBatch();
			                    } else if (ii == 5) {
			                      psSeriesDetl.setBigDecimal(2, new BigDecimal(ii));
			                      psSeriesDetl.setBigDecimal(3, new BigDecimal(unit5));
			                      psSeriesDetl.addBatch();
			                    }
			                }
			              }
			            }
			            else {
			              detlPs.setString(19, "--");
			              detlPs.setString(20, "--");
			            }
			            detlPs.setTimestamp(21, new Timestamp(dateTime.getTime()));
			            detlPs.setBigDecimal(22, null);
			            detlPs.setString(23, divCode);

			            detlPs.addBatch();
			            count++;
			          }
			        }
			        if (count > 0) {
			          String wos = workOrderNumber.split("-").length > 4 ? workOrderNumber.substring(0, workOrderNumber.length() - workOrderNumber.split("-")[4].length() - 1) : workOrderNumber;
			          Long repeatWOLayer = validateRepeatWO(wos, layer, globalDateTime, divCode);
			          if (repeatWOLayer.longValue() == 0L) {
			            psMast.executeBatch();
			            detlPs.executeBatch();
			          }
			          ips.addBatch();
			          if (executeFlag.booleanValue()) {
			            psSeriesMast.executeBatch();
			            psSeriesDetl.executeBatch();
			          }
			        }
			        if (psMastSeriesPkey != null) psMastSeriesPkey.close();
			        if (psSeriesMast != null) psSeriesMast.close();
			        if (psSeriesDetl != null) psSeriesDetl.close();
			        if (rsSeries != null) rsSeries.close();
			        if (rs != null) rs.close();
			        if (psMast != null) psMast.close();
			        if (detlPs != null) detlPs.close();
			        if (conn != null) conn.close(); 
			      }
			    }
			    catch (Exception e) {
			      if (psMastSeriesPkey != null) psMastSeriesPkey.close();
			      if (psSeriesMast != null) psSeriesMast.close();
			      if (psSeriesDetl != null) psSeriesDetl.close();
			      if (rsSeries != null) rsSeries.close();
			      if (rs != null) rs.close();
			      if (psMast != null) psMast.close();
			      if (detlPs != null) detlPs.close();
			      if (conn != null) conn.close();
			      e.printStackTrace();
			      SendError(e.getMessage());
			     
			    } finally {
			        if (psMastSeriesPkey == null) psMastSeriesPkey.close();			   
				    if (psSeriesMast != null) psSeriesMast.close();
				    if (psSeriesDetl != null) psSeriesDetl.close();
				    if (rsSeries != null) rsSeries.close();
				    if (rs != null) rs.close();
				    if (psMast != null) psMast.close();
				    if (detlPs != null) detlPs.close();
				    if (conn != null) conn.close();
			    return count;
			  }
		}
		 
			//查询12小时频段内是否有重复的工单号和对应的层次号
		private Long validateRepeatWO(String wo, String layer, java.util.Date inDate, String divCode) throws Exception  {
		    Long count = Long.valueOf(0L);
		    Session session = DBOptionUtil.getSession();
		    SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy HH:mm");
		    SimpleDateFormat sddf = new SimpleDateFormat("MM/dd/yyyy");
		    java.util.Date currDate = inDate != null ? inDate : new java.util.Date();
		    java.util.Date startDate = null;
		    java.util.Date endDate = null;
		    java.util.Date d1 = sdf.parse(sddf.format(inDate != null ? inDate : new java.util.Date()) + " 08:00");
		    java.util.Date d2 = sdf.parse(sddf.format(inDate != null ? inDate : new java.util.Date()) + " 20:00");
		    Date d3 = new Date(d1.getTime() + 1L * 24L * 60L * 60L * 1000L);
		    if ((currDate.after(d1)) && (currDate.before(d2))) {
		      startDate = d1;
		      endDate = d2;
		    } else if ((currDate.after(d2)) && (currDate.before(d3))) {
		      startDate = d2;
		      endDate = d3;
		    }
		    StringBuffer sql = new StringBuffer();
		    sql.append(" SELECT COUNT(WO) FROM MTG_IPQC_TRAC_MAST ");
		    sql.append(" WHERE TDATE>=:fromDate ");
		    sql.append(" AND TDATE<=:toDate ");
		    sql.append(" AND WO =:workOrder ");
		    sql.append(" AND LAYER=:layer ");
		    sql.append(" AND MTG_IPQC_TRAC_MAST.DIV_CODE =:divCode ");

		    SQLQuery query = session.createSQLQuery(sql.toString());
		    query.setTimestamp("fromDate", startDate);
		    query.setTimestamp("toDate", endDate);
		    query.setString("workOrder", wo.substring(0, 15));
		    query.setString("layer", layer);
		    query.setString("divCode", divCode);

		    count = Long.valueOf(Long.parseLong(query.uniqueResult().toString()));
		    DBOptionUtil.closeSession();
		    return count;
	  }
		
	 public String getCustPartNumber(String wo, String divCode) throws Exception {
		    Session session = DBOptionUtil.getSession();
		    StringBuffer pnSql = new StringBuffer();
		    pnSql.append(" SELECT CUSTOMER_PART_NUMBER FROM (SELECT CUSTOMER_PART_NUMBER FROM DATA0006,DATA0050 ");
		    pnSql.append(" where WORK_ORDER_NUMBER LIKE :wo ");
		    pnSql.append(" AND DATA0006.CUST_PART_PTR = DATA0050.pkey ");
		    pnSql.append(" AND DATA0006.SUB_LEVELS = 1 ");
		    pnSql.append(" AND DATA0006.DIV_CODE =:divCode ");
		    pnSql.append(" ORDER BY DATA0006.PKEY DESC) WHERE ROWNUM = 1 ");
		    Query query = session.createSQLQuery(pnSql.toString());
		    query.setString("wo", "%" + wo + "%");
		    query.setString("divCode", divCode);
		    String pn = (String)query.uniqueResult();
		    DBOptionUtil.closeSession();
		    return pn;
	}		
				  
	  public List<Object[]> searchSpcControlInfo(String pn, String layer, String floor, String divCode) throws Exception {
		    Session session = DBOptionUtil.getSession();
		    StringBuffer sql = new StringBuffer();
		    sql.append("SELECT MTG_SPC_CNTR_DETL.CONTROL_MIN, ");
		    sql.append("MTG_SPC_CNTR_DETL.CONTROL_MAX, ");
		    sql.append("MTG_SPC_CNTR_DETL.LIMIT_MAX ");
		    sql.append("FROM MTG_SPC_CNTR_MAST,MTG_SPC_CNTR_DETL ");
		    sql.append("WHERE MTG_SPC_CNTR_MAST.PKEY = MTG_SPC_CNTR_DETL.MTG_SPC_CNTR_MAST_PTR ");
		    sql.append("AND PN =:pn ");
		    sql.append("AND LAYER =:layer ");
		    sql.append("AND MTG_SPC_CNTR_DETL.dept like:floor ");
		    sql.append("AND MTG_SPC_CNTR_MAST.DIV_CODE =:divCode ");
		    sql.append("AND ROWNUM =1  ");
		    Query query = session.createSQLQuery(sql.toString());
		    query.setString("pn", pn.substring(0, pn.length() - 2));
		    query.setString("layer", layer);
		    query.setString("floor", "%" + floor + "%");
		    query.setString("divCode", divCode);
		    List list = query.list();
		    DBOptionUtil.closeSession();
		    return list;
	  }

	  public BigDecimal getLimitValue(BigDecimal[] array) {
		    BigDecimal[] A = array;
		    BigDecimal max;
		    BigDecimal min = max = A[0];
	
		    for (int i = 0; i < A.length; i++) {
		      if (A[i].compareTo(max) > 0)
		        max = A[i];
		      if (A[i].compareTo(min) < 0) {
		        min = A[i];
		      }
		    }
	
		    return max.subtract(min);
	 }
	  
	  @SuppressWarnings("rawtypes")
	  public int setFQEPQADESDTXTParam(File sourceFile, String monitorId, int fDatePOS, String divCode)
			    throws Exception{
			    int count = 0;
			    int i = 0;
			    String groupIndex = null;
			    String pn = null;
			    String wo = null;
			    String nominalOhms = null;
			    String impedanceName = null;
			    String tolPlus = null;
			    String minusPlus = null;
			    String result = null;
			    String globalResult = "PASS";
			    String average = null;
			    String optDate = null;
			    String optTime = null;
			    //HashMap desdMap = new HashMap();
			    Map<BigDecimal,String> desdMap=new HashMap<BigDecimal,String>(); 
			    
			    String encodeType = FileOptionUtil.getTXTEncodeType(sourceFile);

			    SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy HH:mm");
			    SimpleDateFormat sddf = new SimpleDateFormat("MM/dd/yyyy");
			    SimpleDateFormat sddfHH = new SimpleDateFormat("HH:mm");

			    Connection conn = null;
			    PreparedStatement psMast = null;
			    PreparedStatement detlPs = null;
			    PreparedStatement psMastPkey = null;
			    ResultSet rs = null;
			    conn = DBOptionUtil.getConnection();

			    BigDecimal mastPkey = null;
			    StringBuffer detlSql = new StringBuffer();

			    String sqlMast = "INSERT INTO MTG_ODF_DESD_TEST_MAST(PKEY,WO,PN,CHECK_FLAG,TDATE,CURRENT_VERSION,DIV_CODE) VALUES(:PKEY,:WO,:PN,:CHECK_FLAG,:TDATE,0,'GME')";

			    detlSql.append("INSERT INTO MTG_ODF_DESD_TEST_DETL(PKEY,MTG_ODF_DESD_TEST_MAST_PTR,DESCRIPTION,NOMINAL_OHMS,TOL_PLUS,MINUS_PLUS,AVERAGE,AUDIT_RESULT,TDATE,CURRENT_VERSION,DIV_CODE) ");
			    detlSql.append(" VALUES(MTG_ODF_DESD_TEST_DETL_SEQ.NEXTVAL,:MTG_ODF_DESD_TEST_MAST_PTR,:DESCRIPTION,:NOMINAL_OHMS,:TOL_PLUS,:MINUS_PLUS,:AVERAGE,:AUDIT_RESULT,:TDATE,0,'GME')");

			    psMast = conn.prepareStatement(sqlMast);
			    detlPs = conn.prepareStatement(detlSql.toString());
			    try{
			      String[] lines = FileOptionUtil.readLines(sourceFile, encodeType);
			      int j = 0;
			      for (String line : lines) {
			        count++;
			        String localPn = null;
			        String localWo = null;
			        if (count >= 3) {
			          j++;
			          String[] items = line.toString().split("\\|");
			          if ((items != null) && (items.length > 0)) {
			            i++;
			            groupIndex = items[0];
			            localPn = items[2];
			            if (pn == null) {
			              pn = localPn;
			            }
			            localWo = items[24];
			            if (wo == null) {
			              wo = localWo;
			            }
			            nominalOhms = items[3];
			            impedanceName = items[4];
			            tolPlus = items[5];
			            minusPlus = items[6];
			            result = items[16];
			            if (!"1".equals(result)) {
			              globalResult = "NG";
			            }
			            average = items[17];
			            optDate = items[28];
			            optTime = items[29];

			            if ((wo != null) && (
			              (i == 1) || (!wo.equals(localWo)))) {
			              psMastPkey = conn.prepareStatement("SELECT MTG_ODF_DESD_TEST_MAST_SEQ.NEXTVAL FROM DUAL ");
			              rs = psMastPkey.executeQuery();
			              rs.next();
			              mastPkey = rs.getBigDecimal("NEXTVAL");
			              psMast.setBigDecimal(1, mastPkey);
			              wo = localWo;
			            }

			            if ((lines.length - 2 == j) || (!wo.equals(localWo))) {
			              psMast.setString(2, wo);
			              psMast.setString(3, pn);
			              psMast.setString(4, globalResult);
			              psMast.setTimestamp(5, new Timestamp(sdf.parse(optDate + " " + optTime).getTime()));
			              psMast.addBatch();
			              desdMap.put(mastPkey, globalResult);
			              globalResult = "PASS";
			              wo = localWo;
			            }
			            detlPs.setBigDecimal(1, mastPkey);
			            detlPs.setString(2, impedanceName);
			            detlPs.setBigDecimal(3, nominalOhms != null ? new BigDecimal(nominalOhms) : null);
			            detlPs.setBigDecimal(4, tolPlus != null ? new BigDecimal(tolPlus) : null);
			            detlPs.setBigDecimal(5, minusPlus != null ? new BigDecimal(minusPlus) : null);
			            detlPs.setBigDecimal(6, average != null ? new BigDecimal(average) : null);
			            detlPs.setString(7, "1".equals(result) ? "PASS" : "NG");

			            detlPs.setTimestamp(8, new Timestamp(sdf.parse(optDate + " " + optTime).getTime()));
			            detlPs.addBatch();
			            count++;

			            //System.out.println(items[0] + " " + pn + " " + wo + " " + nominalOhms + " " + impedanceName + " " + tolPlus + " " + minusPlus + " " + result + " " + average + " " + optDate + " " + optTime);
			          }
			        }
			      }
			      if (count > 0) {
			          psMast.executeBatch();
			          if ((desdMap != null) && (!desdMap.isEmpty()) && (desdMap.size() > 0)) {
			            for (Map.Entry m : desdMap.entrySet()) {
			              saveOrUpdateDesdMast(null, wo, (BigDecimal)m.getKey(), null, (String)m.getValue(), "D", "GME");
			            }
			          }
			          detlPs.executeBatch();
			      }
			    } catch (Exception e) {
				      if (desdMap != null) desdMap.clear(); desdMap = null;
				      if (rs != null) rs.close();
				      if (psMast != null) psMast.close();
				      if (detlPs != null) detlPs.close();
				      if (conn != null) conn.close();
				      e.printStackTrace();
				      SendError(e.getMessage());
			  
			    } finally {
				        if (desdMap == null)desdMap.clear();  desdMap = null;
					    if (rs != null) rs.close();
					    if (psMast != null) psMast.close();
					    if (detlPs != null) detlPs.close();
					    if (conn != null) conn.close();
					    return count;
			    }
		}

	  public int setQEPQAFADesdLineWParam(String monitorId, Sheet sheet, int fDatePOS, String divCode)   throws Exception
			  {
			    int count = 0;
			    Connection conn = null;
			    PreparedStatement psMast = null;
			    PreparedStatement detlPs = null;
			    PreparedStatement psMastPkey = null;
			    ResultSet rs = null;
			    @SuppressWarnings("rawtypes")
				HashMap hashColumnMap = new HashMap();
			    conn = DBOptionUtil.getConnection();

			    BigDecimal mastPkey = null;
			    StringBuffer detlSql = new StringBuffer();

			    String sqlMast = "INSERT INTO MTG_ODF_LINE_W_MAST(PKEY,PN,CHECK_FLAG,TDATE,CURRENT_VERSION,DIV_CODE) VALUES(:PKEY,:PN,:CHECK_FLAG,:TDATE,0,'GME')";

			    detlSql.append("INSERT INTO MTG_ODF_LINE_W_DETL(PKEY,MTG_ODF_LINE_W_MAST_PTR,WO,LAYER,TTYPE,LINE_WIDTH,UNIT_1,UNIT_2,UNIT_3,UNIT_4,UNIT_5,AUDIT_RESULT,TDATE,CURRENT_VERSION,DIV_CODE) ");
			    detlSql.append(" VALUES(MTG_ODF_LINE_W_DETL_SEQ.NEXTVAL,:MTG_ODF_LINE_W_MAST_PTR,:WO,:LAYER,:TTYPE,:LINE_WIDTH,:UNIT_1,:UNIT_2,:UNIT_3,:UNIT_4,:UNIT_5,:AUDIT_RESULT,:TDATE,0,'GME')");

			    psMast = conn.prepareStatement(sqlMast);
			    detlPs = conn.prepareStatement(detlSql.toString());

			    psMastPkey = conn.prepareStatement("SELECT MTG_ODF_LINE_W_MAST_SEQ.NEXTVAL FROM DUAL ");
			    rs = psMastPkey.executeQuery();
			    rs.next();
			    mastPkey = rs.getBigDecimal("NEXTVAL");

			    psMast.setBigDecimal(1, mastPkey);
			    try
			    {
			      Row rowm = null;
			      Row rowDataNum = null;
			      Row rowPnNum = null;
			      Cell cell = null;
			      Cell cellData = null;
			      String machNO = monitorId;
			      String pn = null;
			      String wo = null;
			      java.util.Date optDate = null;
			      String globalCheckFlag = "PASS";

			      rowPnNum = sheet.getRow(0);
			      rowm = sheet.getRow(2);
			      rowDataNum = sheet.getRow(3);

			      pn = rowPnNum.getCell(1).getStringCellValue();

			      psMast.setString(2, (pn != null) && (pn.indexOf("-") > 0) ? pn.split("-")[0] : pn);

			      for (int i = 0; i <= 256; i += 7) {
			        cell = rowm.getCell(i + 1);
			        cellData = rowDataNum.getCell(i + 1);
			        if (cellData != null) {
			          hashColumnMap.put(Integer.valueOf(cell.getColumnIndex()), cell.getStringCellValue());
			        }
			      }
			      if ((hashColumnMap != null) && (!hashColumnMap.isEmpty()) && (hashColumnMap.size() > 0)) {
			        List arrayList = new ArrayList(hashColumnMap.entrySet());
			        Collections.sort(arrayList, new Comparator() {
			          public int compare(Object o1, Object o2) {
			            Map.Entry obj1 = (Map.Entry)o1;
			            Map.Entry obj2 = (Map.Entry)o2;
			            return ((Integer)obj1.getKey()).compareTo((Integer)obj2.getKey());
			          }
			        });
			        String layer = "";
			        String ttype = "";
			        String lineWidth = null;
			        String actualValue1 = null;
			        String actualValue2 = null;
			        String actualValue3 = null;
			        String actualValue4 = null;
			        String actualValue5 = null;
			        String checkFlag = "PASS";
			        int rowNum = sheet.getLastRowNum();
			        for (Iterator iter = arrayList.iterator(); iter.hasNext(); ) {
			          Map.Entry entry = (Map.Entry)iter.next();
			          Integer key = (Integer)entry.getKey();
			          if (entry.getValue() != null) {
			            layer = entry.getValue().toString().split(" ").length > 1 ? entry.getValue().toString().replaceAll(" +", "@").split("@")[0] : "";
			            ttype = entry.getValue().toString().split(" ").length > 1 ? entry.getValue().toString().replaceAll(" +", "@").split("@")[1] : "";
			            int j = 1;
			            for (int i = 3; i <= rowNum; i++) {
			              int n = 1;
			              for (int k = Integer.parseInt(entry.getKey().toString()) == 1 ? Integer.parseInt(entry.getKey().toString()) : Integer.parseInt(entry.getKey().toString()) + 1; k <= (Integer.parseInt(entry.getKey().toString()) == 1 ? Integer.parseInt(entry.getKey().toString()) : Integer.parseInt(entry.getKey().toString()) + 1) + 4; k++) {
			                rowm = sheet.getRow(i);
			                if (rowm != null) {
			                  cell = rowm.getCell(k);
			                  if (cell != null) {
			                    if ((n == 1) && (cell != null) && (cell.getCellType() == 1) && 
			                      (cell.getStringCellValue() != null)) {
			                      lineWidth = cell.getStringCellValue();
			                    }

			                    if ((n == 2) && (cell != null) && (cell.getCellType() == 0)) {
			                      if (j == 1)
			                        actualValue1 = String.valueOf(cell.getNumericCellValue());
			                      else if (j == 2)
			                        actualValue2 = String.valueOf(cell.getNumericCellValue());
			                      else if (j == 3)
			                        actualValue3 = String.valueOf(cell.getNumericCellValue());
			                      else if (j == 4)
			                        actualValue4 = String.valueOf(cell.getNumericCellValue());
			                      else if (j == 5) {
			                        actualValue5 = String.valueOf(cell.getNumericCellValue());
			                      }
			                    }
			                    if ((n == 3) && (cell != null) && (cell.getCellType() == 0)) {
			                      if (cell.getNumericCellValue() != 0.0D) {
			                        checkFlag = "NG";
			                        globalCheckFlag = "NG";
			                      } else {
			                        checkFlag = "PASS";
			                      }
			                    }
			                    if ((n == 4) && (cell != null) && (cell.getCellType() == 1) && 
			                      (cell.getStringCellValue() != null)) {
			                      wo = cell.getStringCellValue();
			                    }

			                    if ((n == 5) && (cell != null) && (cell.getCellType() == 0)) {
			                      optDate = cell.getDateCellValue();
			                    }
			                  }

			                  n++;
			                  if (n > 5) {
			                    n = 1;
			                  }
			                }
			              }

			              j++;
			              if (j > 5) {
			                j = 1;
			                detlPs.setBigDecimal(1, mastPkey);
			                detlPs.setString(2, wo);
			                detlPs.setString(3, layer);
			                detlPs.setString(4, ttype);
			                detlPs.setString(5, lineWidth);
			                detlPs.setBigDecimal(6, actualValue1 != null ? new BigDecimal(actualValue1) : null);
			                detlPs.setBigDecimal(7, actualValue2 != null ? new BigDecimal(actualValue2) : null);
			                detlPs.setBigDecimal(8, actualValue3 != null ? new BigDecimal(actualValue3) : null);
			                detlPs.setBigDecimal(9, actualValue4 != null ? new BigDecimal(actualValue4) : null);
			                detlPs.setBigDecimal(10, actualValue5 != null ? new BigDecimal(actualValue5) : null);
			                detlPs.setString(11, checkFlag);
			                detlPs.setTimestamp(12, new Timestamp(optDate.getTime()));

			                detlPs.addBatch();
			                count++;
			                System.out.println(pn + " " + layer + " " + ttype + " " + lineWidth + " " + actualValue1 + " " + actualValue2 + " " + actualValue3 + " " + actualValue4 + " " + actualValue5 + " " + checkFlag + " " + wo + " " + optDate);
			              }
			            }
			            psMast.setString(3, globalCheckFlag);
			            psMast.setTimestamp(4, new Timestamp(optDate.getTime()));
			          }
			        }
			      }
			      if (count > 0) {
			        psMast.addBatch();
			        psMast.executeBatch();
			        detlPs.executeBatch();
			      }
			    } catch (Exception e) {
			      if (rs != null) rs.close();
			      if (psMast != null) psMast.close();
			      if (detlPs != null) detlPs.close();
			      if (conn != null) conn.close();
			      e.printStackTrace();
			      SendError(e.getMessage());
			      System.out.println(e.getMessage());
			    } finally {
			      hashColumnMap.clear();
			    }hashColumnMap = null;
			    if (rs != null) rs.close();
			    if (psMast != null) psMast.close();
			    if (detlPs != null) detlPs.close();
			    if (conn != null) conn.close();
			    return count;
	}

	 public int setQEPQAIPQCDesdLineWParam(String monitorId, Sheet sheet, int fDatePOS, String divCode) throws Exception {
			    int count = 0;
			    Connection conn = null;
			    PreparedStatement psMast = null;
			    PreparedStatement detlPs = null;
			    PreparedStatement psMastPkey = null;
			    ResultSet rs = null;
			    Map hashColumnMap = new HashMap();
			    Map<BigDecimal,String> desdMap=new HashMap<BigDecimal,String>(); 
			    conn = DBOptionUtil.getConnection();
			
			    BigDecimal mastPkey = null;
			    StringBuffer detlSql = new StringBuffer();

			    String sqlMast = "INSERT INTO MTG_ODF_LINE_W_MAST(PKEY,PN,CHECK_FLAG,TDATE,CURRENT_VERSION,DIV_CODE) VALUES(:PKEY,:PN,:CHECK_FLAG,:TDATE,0,'GME')";

			    detlSql.append("INSERT INTO MTG_ODF_LINE_W_DETL(PKEY,MTG_ODF_LINE_W_MAST_PTR,WO,LAYER,TTYPE,LINE_WIDTH,UNIT_1,UNIT_2,UNIT_3,UNIT_4,UNIT_5,AUDIT_RESULT,TDATE,CURRENT_VERSION,DIV_CODE) ");
			    detlSql.append(" VALUES(MTG_ODF_LINE_W_DETL_SEQ.NEXTVAL,:MTG_ODF_LINE_W_MAST_PTR,:WO,:LAYER,:TTYPE,:LINE_WIDTH,:UNIT_1,:UNIT_2,:UNIT_3,:UNIT_4,:UNIT_5,:AUDIT_RESULT,:TDATE,0,'GME')");

			    psMast = conn.prepareStatement(sqlMast);
			    detlPs = conn.prepareStatement(detlSql.toString());

			    psMastPkey = conn.prepareStatement("SELECT MTG_ODF_LINE_W_MAST_SEQ.NEXTVAL FROM DUAL ");
			    rs = psMastPkey.executeQuery();
			    rs.next();
			    mastPkey = rs.getBigDecimal("NEXTVAL");

			    psMast.setBigDecimal(1, mastPkey);
			    try
			    {
			      Row rowm = null;
			      Row rowDataNum = null;
			      Row rowPnNum = null;
			      Cell cell = null;
			      Cell cellData = null;
			      String machNO = monitorId;
			      String pn = null;
			      String wo = null;
			      java.util.Date optDate = null;
			      String globalCheckFlag = "PASS";

			      rowPnNum = sheet.getRow(0);
			      rowm = sheet.getRow(1);
			      rowDataNum = sheet.getRow(3);

			      pn = rowPnNum.getCell(1).getStringCellValue();

			      psMast.setString(2, (pn != null) && (pn.indexOf("-") > 0) ? pn.split("-")[0] : pn);

			      for (int i = 0; i <= 256; i += 7) {
			        cell = rowm.getCell(i + 1);
			        cellData = rowDataNum.getCell(i + 1);
			        if (cellData != null)
			          hashColumnMap.put(Integer.valueOf(cell.getColumnIndex()), cell.getStringCellValue());
			      }
			      String layer;
			      if ((hashColumnMap != null) && (!hashColumnMap.isEmpty()) && (hashColumnMap.size() > 0)) {
			        List arrayList = new ArrayList(hashColumnMap.entrySet());
			        Collections.sort(arrayList, new Comparator() {
			          public int compare(Object o1, Object o2) {
			            Map.Entry obj1 = (Map.Entry)o1;
			            Map.Entry obj2 = (Map.Entry)o2;
			            return ((Integer)obj1.getKey()).compareTo((Integer)obj2.getKey());
			          }
			        });
			        layer = "";
			        String ttype = "";
			        String lineWidth = null;
			        String actualValue1 = null;
			        String actualValue2 = null;
			        String actualValue3 = null;
			        String actualValue4 = null;
			        String actualValue5 = null;
			        String checkFlag = "PASS";
			        int rowNum = sheet.getLastRowNum();
			        for (Iterator iter = arrayList.iterator(); iter.hasNext(); ) {
			          Map.Entry entry = (Map.Entry)iter.next();
			          Integer key = (Integer)entry.getKey();
			          if (entry.getValue() != null) {
			            layer = entry.getValue().toString().split(" ").length > 1 ? entry.getValue().toString().replaceAll(" +", "@").split("@")[0] : "";
			            ttype = entry.getValue().toString().split(" ").length > 1 ? entry.getValue().toString().replaceAll(" +", "@").split("@")[1] : "";
			            int j = 1;
			            for (int i = 3; i <= rowNum; i++) {
			              int n = 1;
			              for (int k = Integer.parseInt(entry.getKey().toString()) == 1 ? Integer.parseInt(entry.getKey().toString()) : Integer.parseInt(entry.getKey().toString()); k <= (Integer.parseInt(entry.getKey().toString()) == 1 ? Integer.parseInt(entry.getKey().toString()) : Integer.parseInt(entry.getKey().toString())) + 4; k++) {
			                rowm = sheet.getRow(i);
			                if (rowm != null) {
			                  cell = rowm.getCell(k);
			                  if (cell != null) {
			                    if ((n == 1) && (cell != null) && (cell.getCellType() == 1) && 
			                      (cell.getStringCellValue() != null)) {
			                      lineWidth = cell.getStringCellValue();
			                    }

			                    if ((n == 2) && (cell != null) && (cell.getCellType() == 0)) {
			                      if (j == 1)
			                        actualValue1 = String.valueOf(cell.getNumericCellValue());
			                      else if (j == 2)
			                        actualValue2 = String.valueOf(cell.getNumericCellValue());
			                      else if (j == 3)
			                        actualValue3 = String.valueOf(cell.getNumericCellValue());
			                      else if (j == 4)
			                        actualValue4 = String.valueOf(cell.getNumericCellValue());
			                      else if (j == 5) {
			                        actualValue5 = String.valueOf(cell.getNumericCellValue());
			                      }
			                    }
			                    if ((n == 3) && (cell != null) && (cell.getCellType() == 0)) {
			                      if (cell.getNumericCellValue() != 0.0D) {
			                        checkFlag = "NG";
			                        globalCheckFlag = "NG";
			                      } else {
			                        checkFlag = "PASS";
			                      }
			                    }
			                    if ((n == 4) && (cell != null) && (cell.getCellType() == 1) && 
			                      (cell.getStringCellValue() != null)) {
			                      wo = cell.getStringCellValue();
			                    }

			                    if ((n == 5) && (cell != null) && (cell.getCellType() == 0)) {
			                      optDate = cell.getDateCellValue();
			                    }
			                  }

			                  n++;
			                  if (n > 5) {
			                    n = 1;
			                  }
			                }
			              }

			              j++;
			              if (j > 5) {
			                j = 1;
			                detlPs.setBigDecimal(1, mastPkey);
			                detlPs.setString(2, wo);
			                detlPs.setString(3, layer);
			                detlPs.setString(4, ttype);
			                detlPs.setString(5, lineWidth);
			                detlPs.setBigDecimal(6, actualValue1 != null ? new BigDecimal(actualValue1) : null);
			                detlPs.setBigDecimal(7, actualValue2 != null ? new BigDecimal(actualValue2) : null);
			                detlPs.setBigDecimal(8, actualValue3 != null ? new BigDecimal(actualValue3) : null);
			                detlPs.setBigDecimal(9, actualValue4 != null ? new BigDecimal(actualValue4) : null);
			                detlPs.setBigDecimal(10, actualValue5 != null ? new BigDecimal(actualValue5) : null);
			                detlPs.setString(11, checkFlag);
			                detlPs.setTimestamp(12, new Timestamp(optDate.getTime()));

			                detlPs.addBatch();
			                count++;
			                System.out.println(pn + " " + layer + " " + ttype + " " + lineWidth + " " + actualValue1 + " " + actualValue2 + " " + actualValue3 + " " + actualValue4 + " " + actualValue5 + " " + checkFlag + " " + wo + " " + optDate);
			              }
			            }
			            psMast.setString(3, globalCheckFlag);
			            psMast.setTimestamp(4, new Timestamp(optDate.getTime()));
			            desdMap.put(mastPkey, globalCheckFlag);
			          }
			        }
			      }
			      if (count > 0) {
			        psMast.addBatch();
			        if ((desdMap != null) && (!desdMap.isEmpty()) && (desdMap.size() > 0)) {
			          for (Map.Entry m : desdMap.entrySet()) {
			            saveOrUpdateDesdMast((pn != null) && (pn.indexOf("-") > 0) ? pn.split("-")[0] : pn, null, 
			              (BigDecimal)m.getKey(), (String)m.getValue(), null, "L", "GME");
			          }
			        }
			        psMast.executeBatch();
			        detlPs.executeBatch();
			      }
			    } catch (Exception e) {
			      if (desdMap != null) desdMap.clear(); desdMap = null;
			      if (rs != null) rs.close();
			      if (psMast != null) psMast.close();
			      if (detlPs != null) detlPs.close();
			      if (conn != null) conn.close();
			      e.printStackTrace();
			      SendError(e.getMessage());
			  
			    } finally {
				      if (desdMap == null){  
				    	  desdMap.clear(); 
				    	  desdMap = null;
				      }	
				      hashColumnMap.clear();
				      hashColumnMap = null;
				    if (rs != null) rs.close();
				    if (psMast != null) psMast.close();
				    if (detlPs != null) detlPs.close();
				    if (conn != null) conn.close();
				    return count;
			   }
		 }

		 public int setQEPQA410CDESDParam(String monitorId, Sheet sheet, int fDatePOS, String divCode) throws Exception
			  {
			    int count = 0;
			    Connection conn = null;
			    PreparedStatement psMast = null;
			    PreparedStatement detlPs = null;
			    PreparedStatement psMastPkey = null;
			    ResultSet rs = null;
			    SimpleDateFormat sdfd = new SimpleDateFormat("yyyy-MM-dd HH:mm");
			    Map<BigDecimal,String> desdMap=new HashMap<BigDecimal,String>(); 
			    try {
			      Row rowm = null;
			      Cell cell = null;
			      String machNO = monitorId;

			      conn = DBOptionUtil.getConnection();

			      BigDecimal mastPkey = null;
			      StringBuffer detlSql = new StringBuffer();

			      String sqlMast = "INSERT INTO MTG_ODF_DESD_TEST_MAST(PKEY,WO,PN,CHECK_FLAG,TDATE,CURRENT_VERSION,DIV_CODE) VALUES(:PKEY,:WO,:PN,:CHECK_FLAG,:TDATE,0,'GME')";

			      detlSql.append("INSERT INTO MTG_ODF_DESD_TEST_DETL(PKEY,MTG_ODF_DESD_TEST_MAST_PTR,DESCRIPTION,NOMINAL_OHMS,TOL_PLUS,MINUS_PLUS,AVERAGE,AUDIT_RESULT,TDATE,CURRENT_VERSION,DIV_CODE) ");
			      detlSql.append(" VALUES(MTG_ODF_DESD_TEST_DETL_SEQ.NEXTVAL,:MTG_ODF_DESD_TEST_MAST_PTR,:DESCRIPTION,:NOMINAL_OHMS,:TOL_PLUS,:MINUS_PLUS,:AVERAGE,:AUDIT_RESULT,:TDATE,0,'GME')");

			      psMast = conn.prepareStatement(sqlMast);
			      detlPs = conn.prepareStatement(detlSql.toString());

			      int rowNum = sheet.getLastRowNum();
			      String pn = null;
			      String wo = null;
			      String nominalOhms = null;
			      String impedanceName = null;
			      String tolPlus = null;
			      String minusPlus = null;
			      String result = null;
			      String average = null;
			      java.util.Date optDate = null;
			      java.util.Date optTime = null;
			      String globalResult = "PASS";

			      int k = 0;
			      String localWo;
			      for (int i = 1; i <= rowNum; i++) {
			        rowm = sheet.getRow(i);
			        localWo = null;
			        k++;
			        if (rowm != null) {
			          for (int j = 1; j <= 20; j++) {
			            cell = rowm.getCell(j);
			            if (cell != null) {
			              if ((j == 1) && (cell != null) && (cell.getCellType() == 1)) {
			                impedanceName = cell.getStringCellValue();
			              }
			              if ((j == 2) && (cell != null) && (cell.getCellType() == 0)) {
			                nominalOhms = String.valueOf(cell.getNumericCellValue());
			              }
			              if ((j == 3) && (cell != null) && (cell.getCellType() == 0)) {
			                tolPlus = String.valueOf(cell.getNumericCellValue());
			              }
			              if ((j == 4) && (cell != null) && (cell.getCellType() == 0)) {
			                minusPlus = String.valueOf(cell.getNumericCellValue());
			              }
			              if ((j == 5) && (cell != null) && (cell.getCellType() == 0)) {
			                average = String.valueOf(cell.getNumericCellValue());
			              }
			              if ((j == 8) && (cell != null) && (cell.getCellType() == 1)) {
			                result = cell.getStringCellValue();
			                if (!"PASS".equals(result)) {
			                  globalResult = "NG";
			                }
			              }
			              if ((j == 16) && (cell != null) && (cell.getCellType() == 1)) {
			                localWo = cell.getStringCellValue();
			                if (wo == null) {
			                  wo = localWo;
			                }
			              }
			              if ((j == 18) && (cell != null) && (cell.getCellType() == 0)) {
			                optDate = cell.getDateCellValue();
			              }
			              if ((j == 19) && (cell != null) && (cell.getCellType() == 0)) {
			                optTime = cell.getDateCellValue();
			              }
			            }
			          }

			          if ((wo != null) && (
			            (k == 1) || (!wo.equals(localWo)))) {
			            psMastPkey = conn.prepareStatement("SELECT MTG_ODF_DESD_TEST_MAST_SEQ.NEXTVAL FROM DUAL ");
			            rs = psMastPkey.executeQuery();
			            rs.next();
			            mastPkey = rs.getBigDecimal("NEXTVAL");

			            pn = getCustPartNumber(wo, "GME");

			            psMast.setBigDecimal(1, mastPkey);
			            psMast.setString(2, wo);
			            psMast.setString(3, pn);
			            wo = localWo;
			          }

			          if ((rowNum == k) || (!wo.equals(localWo))) {
			            psMast.setString(2, wo);

			            pn = getCustPartNumber(wo, "GME");

			            psMast.setString(3, pn);
			            psMast.setString(4, globalResult);
			            psMast.setTimestamp(5, new Timestamp(sdfd.parse(new SimpleDateFormat("yyyy-MM-dd").format(optDate) + " " + new SimpleDateFormat("HH:mm").format(optTime)).getTime()));
			            psMast.addBatch();
			            desdMap.put(mastPkey, globalResult);
			            globalResult = "PASS";
			            wo = localWo;
			          }
			          detlPs.setBigDecimal(1, mastPkey);
			          detlPs.setString(2, impedanceName);
			          detlPs.setBigDecimal(3, nominalOhms != null ? new BigDecimal(nominalOhms) : null);
			          detlPs.setBigDecimal(4, tolPlus != null ? new BigDecimal(tolPlus) : null);
			          detlPs.setBigDecimal(5, minusPlus != null ? new BigDecimal(minusPlus) : null);
			          detlPs.setBigDecimal(6, average != null ? new BigDecimal(average) : null);
			          detlPs.setString(7, result);
			          detlPs.setTimestamp(8, new Timestamp(sdfd.parse(new SimpleDateFormat("yyyy-MM-dd").format(optDate) + " " + new SimpleDateFormat("HH:mm").format(optTime)).getTime()));
			          detlPs.addBatch();
			          count++;

			          System.out.println(impedanceName + " " + nominalOhms + " " + tolPlus + " " + minusPlus + " " + average + " " + result + " " + wo + " " + optDate + " " + optTime);
			        }
			      }
			      if (count > 0) {
			        psMast.executeBatch();
			        if ((desdMap != null) && (!desdMap.isEmpty()) && (desdMap.size() > 0)) {
			          for (Map.Entry m : desdMap.entrySet()) {
                           saveOrUpdateDesdMast(null, wo,(BigDecimal)m.getKey(), null, (String)m.getValue(), "D", "GME");
			          }
			        }
			        detlPs.executeBatch();
			      }
			    } catch (Exception e) {
				      if (rs != null) rs.close();
				      if (psMast != null) psMast.close();
				      if (detlPs != null) detlPs.close();
				      if (conn != null) conn.close();
				      e.printStackTrace();
				      SendError(e.getMessage());
			
			    } finally {
			     
			    	if (rs == null)     rs.close();
				    if (psMast != null) psMast.close();
				    if (detlPs != null) detlPs.close();
				    if (conn != null) conn.close();
			         return count;
			 }
	  }	  
	  
	  public String getPNSeries(String pn, String divCode) throws Exception {
		    Session session = DBOptionUtil.getSession();
		    StringBuffer sql = new StringBuffer();
		    sql.append("SELECT FPPRD01 FROM MTG_Product_Series WHERE ");
		    sql.append("FPPRD03 LIKE :pn ");
		    sql.append("AND DIV_CODE =:divCode AND ROWNUM = 1");
		    Query query = session.createSQLQuery(sql.toString());
		    query.setString("pn", pn + "%");
		    query.setString("divCode", divCode);
		    String seriesNo = (String)query.uniqueResult();
		    DBOptionUtil.closeSession();
		    return seriesNo;
	  }  
	  
	  public List<Object[]> getODFDesdFAList(String custPartNumber, String wo, String lineImpdFlag, String divCode) throws Exception {
	    Session session = DBOptionUtil.getSession();
	    StringBuffer sql = new StringBuffer();
	    sql.append(" SELECT PKEY,CUSTOMER_PART_NUMBER,WORK_ORDER_NUMBER,TDATE,CATEGORY ,");
	    sql.append("LINE_W_PTR,IMPD_PTR,LINE_W_RESULT,IMPD_RESULT,LINE_IMPD_RESULT FROM ");
	    sql.append("(SELECT MTG_DESD_TRAC_MAST.PKEY,CUSTOMER_PART_NUMBER,WORK_ORDER_NUMBER,MTG_DESD_TRAC_MAST.CATEGORY,");
	    sql.append("MTG_DESD_TRAC_MAST.TDATE,LINE_W_PTR,IMPD_PTR,LINE_W_RESULT,IMPD_RESULT,LINE_IMPD_RESULT ");
	    sql.append("FROM MTG_DESD_TRAC_MAST,DATA0006,DATA0050 WHERE ");
	    sql.append("MTG_DESD_TRAC_MAST.WO_PTR = DATA0006.PKEY ");
	    sql.append("AND DATA0006.CUST_PART_PTR = DATA0050.PKEY ");
	    if ("L".equals(lineImpdFlag))
	      sql.append("AND DATA0050.CUSTOMER_PART_NUMBER LIKE :pn ");
	    else if ("D".equals(lineImpdFlag)) {
	      sql.append("AND DATA0006.WORK_ORDER_NUMBER LIKE :wo ");
	    }
	    sql.append("AND MTG_DESD_TRAC_MAST.DIV_CODE =:divCode ");
	    sql.append("AND (IMPD_PTR IS NULL OR IMPD_PTR IS NULL OR LINE_IMPD_RESULT IS NULL) ");

	    Query query = session.createSQLQuery(sql.toString());
	    if ("L".equals(lineImpdFlag))
	      query.setString("pn", custPartNumber + "%");
	    else if ("D".equals(lineImpdFlag)) {
	      query.setString("wo", wo + "%");
	    }
	    query.setString("divCode", divCode);
	    List list = query.list();
	    DBOptionUtil.closeSession();
	    return list;
	  }	  
	  
	  public Boolean saveOrUpdateDesdMast(String custPartNumber, String wo, BigDecimal lineDesdPtr, 
			  String lineWResults, String impdResults, String lineImpdFlag, String divCode)
			    throws Exception {
			    List list = getODFDesdFAList(custPartNumber, wo, lineImpdFlag, divCode);
			    Session session = DBOptionUtil.getSession();
			    if ((list != null) && (list.size() > 0)) {
			      for (int i = 0; i < list.size(); i++) {
			        BigDecimal mastPkey = ((Object[])list.get(i))[0] == null ? null : new BigDecimal(((Object[])list.get(0))[0].toString());
			        String lineWResult = ((Object[])list.get(i))[7] == null ? null : ((Object[])list.get(0))[7].toString();
			        String impdResult = ((Object[])list.get(i))[8] == null ? null : ((Object[])list.get(0))[8].toString();
			        String category = ((Object[])list.get(i))[4] == null ? "" : ((Object[])list.get(0))[4].toString();
			        StringBuffer sql = new StringBuffer();

			        if (!"2".equals(category)) {
			          sql.append("UPDATE MTG_DESD_TRAC_MAST SET ");

			          if (((lineWResult == null) || (!"".equals(lineWResults))) && ("L".equals(lineImpdFlag))) {
			            sql.append(" LINE_W_PTR=:lineWPtr,LINE_W_RESULT=:lineWResult ");
			          }
			          if (((impdResult == null) || (!"".equals(impdResults))) && ("D".equals(lineImpdFlag))) {
			            sql.append(" IMPD_PTR=:impdPtr,IMPD_RESULT=:impdResult ");
			          }
			          if ((impdResult != null) && (lineWResults != null) && ("L".equals(lineImpdFlag))) {
			            sql.append(" ,LINE_IMPD_RESULT=:lineImpdResult,STATUS=:status ");
			          }
			          if ((lineWResult != null) && (impdResults != null) && ("D".equals(lineImpdFlag))) {
			            sql.append(" ,LINE_IMPD_RESULT=:lineImpdResult,STATUS=:status ");
			          }
			          sql.append(" WHERE PKEY=:PKEY ");
			          Query query = session.createSQLQuery(sql.toString());
			          if (((lineWResult == null) || (!"".equals(lineWResults))) && ("L".equals(lineImpdFlag))) {
			            query.setBigDecimal("lineWPtr", lineDesdPtr);
			            query.setString("lineWResult", lineWResults);
			          }
			          if (((impdResult == null) || (!"".equals(impdResults))) && ("D".equals(lineImpdFlag))) {
			            query.setBigDecimal("impdPtr", lineDesdPtr);
			            query.setString("impdResult", impdResults);
			          }
			          if ((impdResult != null) && (lineWResults != null) && ("L".equals(lineImpdFlag))) {
			            if (("PASS".equals(impdResult)) && ("PASS".equals(lineWResults))) {
			              query.setString("lineImpdResult", "PASS");
			              query.setInteger("status", 0);
			            } else {
			              query.setString("lineImpdResult", "NG");
			              query.setInteger("status", 1);
			            }
			          }
			          if ((lineWResult != null) && (impdResults != null) && ("D".equals(lineImpdFlag))) {
			            if (("PASS".equals(impdResults)) && ("PASS".equals(lineWResult))) {
			              query.setString("lineImpdResult", "PASS");
			              query.setInteger("status", 0);
			            } else {
			              query.setString("lineImpdResult", "NG");
			              query.setInteger("status", 1);
			            }
			          }
			          query.setBigDecimal("PKEY", mastPkey);
			          query.executeUpdate();
			          sql = null;
			        }
			      }
			    }
			    DBOptionUtil.closeSession();
			    return Boolean.valueOf(true);
			  }	  
	  

}
