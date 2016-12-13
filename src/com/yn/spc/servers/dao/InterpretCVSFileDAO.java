package com.yn.spc.servers.dao;

import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import com.yn.spc.util.CsvUtil;
import com.yn.spc.util.OutRunTimeLog;



public class InterpretCVSFileDAO {
    
	@SuppressWarnings("finally")
	public int CommonCSVFileParsing(PreparedStatement ips, String monitorId,CsvUtil csv,int fDatePOS,String divCode) throws SQLException, ParseException{
		int count=0,pos,inc =0 ;
		int rowNum = csv.getRowNum(); 
		int colNum = csv.getColNum();
		String cell;
		try{
			for (int i = 1; i < rowNum; i++) {	
				
				if("P-LDR-Effic-01".equals(monitorId)){
				    cell = csv.filename;
				    pos = cell.indexOf("#");
				    if(pos >=1 ){
				    	cell = cell.substring(0, pos+1);				    	
				    	pos = cell.indexOf("_");				    	
				    	if (pos >=1 )
				    		cell = cell.substring(pos+1);
				    }else{
				    	cell = "0#" ; 
				    }
				    inc =1 ;
				    ips.setString(inc, cell.trim());				  
				}
				
			    for(int j=0; j< colNum; j++){			    	
			    	cell = csv.getString(i,j);			    	
			    	if(fDatePOS > 0 && (((inc ==1) && (fDatePOS-j==2)) ||  (fDatePOS-j)==1)){ 
			    		if(cell == null || cell.equals("")){		
			    			
			    		   ips.setDate(j+1+inc, new java.sql.Date(new Date().getTime()));
			    		   
						}else {
							// 2013/09/10 
							if(cell.startsWith("'")){
								cell = cell.replace("'", "");
							}
							cell = cell.trim();
						
							if(cell.substring(4, 4).equals("-")){
							   ips.setDate(j+1+inc, new java.sql.Date(new SimpleDateFormat("yyyy-MM-dd").parse(cell).getTime()));
							}else if(cell.substring(4, 4).equals("/")){
								ips.setDate(j+1+inc, new java.sql.Date(new SimpleDateFormat("yyyy/MM/dd").parse(cell).getTime()));
							}else if(cell.substring(2, 2).equals("-")){
								ips.setDate(j+1+inc, new java.sql.Date(new SimpleDateFormat("MM-dd-yyyy").parse(cell).getTime()));
							}else if(cell.substring(2, 2).equals("/")){
								ips.setDate(j+1+inc, new java.sql.Date(new SimpleDateFormat("MM/dd/yyyy").parse(cell).getTime()));
							}else{
								ips.setDate(j+1+inc, new java.sql.Date(new Date().getTime()));
							}								
						}
			    	}else{
			    		if(cell==null ){
			    			ips.setString(j+1+inc,"");
			    		}else{
			    			cell = cell.replace("'", "''");
			    			cell = cell.replace(",", "、");
			    			cell = cell.replace("&", "'|| CHR(38) ||'");			
				        	ips.setString(j+1+inc, cell.trim());	
				        }		    	  		    	   
			    	}			    	
			    }
			 
				if(fDatePOS == 0){			
					ips.setDate(colNum+1+inc, new java.sql.Date(new Date().getTime()));
					ips.setString(colNum+2+inc, divCode);
				}else{
					ips.setString(colNum+1+inc, divCode);				
				}
			    ips.addBatch();
			    
			    count = count + 1;
			    
			}
		}catch (Exception e){
			e.printStackTrace();
			 OutRunTimeLog.showLog(e.getMessage());
		}finally{
			return count++;
		}
	}	
}
