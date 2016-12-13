package com.yn.spc.util;

import java.io.FileNotFoundException;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;

public class Log {  
	 private static Logger loger;   
	 
	 private static Logger newInstance() throws FileNotFoundException{
		  //获得日志类loger的实例  
		  loger=Logger.getLogger(Log.class);  
		  //loger所需的配置文件路径 
		  PropertyConfigurator.configure(System.getProperty("user.dir") + "/conf/log4j.properties");  
		  return loger;
	 }
	   
	 public static Logger getLoger() throws FileNotFoundException{  
	  if(loger!=null)  
	   return loger; 
	  else  
	   return newInstance();  
	 } 
}  