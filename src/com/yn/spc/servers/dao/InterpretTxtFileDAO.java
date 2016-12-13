package com.yn.spc.servers.dao;

import java.io.File;
import java.io.IOException;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.regex.Pattern;

import com.yn.spc.common.Keys;
import com.yn.spc.util.FileOptionUtil;
import com.yn.spc.util.OutRunTimeLog;
import com.yn.spc.util.ParseUtil;



public class InterpretTxtFileDAO {
	   /*
	    * 解析文本文件
	    */
		@SuppressWarnings("finally")
		public int setCommonTXTParam(PreparedStatement ips, File sourceFile, String monitorId,int fDatePOS,String divCode) throws Exception{
			int count=0,colNum=0;
			Pattern pattern03 = Pattern.compile("\t{1}");
			String[] patternStr = null;	    
			String encodeType=FileOptionUtil.getTXTEncodeType(sourceFile);
			String[] lines = FileOptionUtil.readLines(sourceFile,encodeType);
			try{
				for (String line : lines) {
					if(count==0){            //忽略第一行
						count++;
						continue;
					}
					patternStr = pattern03.split(line);
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
				//System.out.println(e.getMessage());
			}finally{
				return count;
			}	
		}
		
		@SuppressWarnings("finally")
		public int ImpFischerXData(PreparedStatement ips, File sourceFile, String monitorId,int fDatePOS,String divCode) throws Exception{
			int count=0,colNum=0;
			try{
				Pattern pattern03 = Pattern.compile("\t{1}");
				String[] patternStr = null;	    
				String encodeType=FileOptionUtil.getTXTEncodeType(sourceFile);
				String[] lines = FileOptionUtil.readLines(sourceFile,encodeType);
				SimpleDateFormat sd  =   new SimpleDateFormat(Keys.DATE_FORMAT5);
				for (String line : lines) {	
					
					while(true){
						 if(line.indexOf("		") ==-1){
							 break;
						 }else{
					         line = line.replaceAll("		", "	");
					     }
					}
					patternStr = pattern03.split(line);
					for(colNum=0; colNum<patternStr.length; colNum++){
						if(fDatePOS > 0 && (fDatePOS-colNum)==1){ 
							if("".equals( ParseUtil.getObjStr(patternStr[colNum]))){
							    ips.setDate(colNum+1, new java.sql.Date(new Date().getTime()));
							}else{
							    ips.setDate(colNum+1,new java.sql.Date(sd.parse(ParseUtil.getObjStr(patternStr[colNum])).getTime()));
							}  
						}else{
							if(colNum == 2){
								ips.setString(colNum+1, ParseUtil.getObjStr(patternStr[4]).trim());
							}else if(colNum == 3){
								ips.setString(colNum+1, ParseUtil.getObjStr(patternStr[2]).trim());
							}else if(colNum == 4){
								ips.setString(colNum+1, ParseUtil.getObjStr(patternStr[5]).trim());
							}else if(colNum == 5){
								ips.setString(colNum+1, ParseUtil.getObjStr(patternStr[3]).trim());
							}else{
							    ips.setString(colNum+1, ParseUtil.getObjStr(patternStr[colNum]).trim());
							}
						}					
					}
					if(fDatePOS == 0){			
						ips.setDate(colNum+1, new java.sql.Date(new Date().getTime()));
						ips.setString(colNum+2, "");
						ips.setString(colNum+3, divCode);
					}else{
						ips.setString(colNum+1, "");	
						ips.setString(colNum+2, divCode);				
					}
					ips.addBatch();
					count=count+1;
				}			
			}catch (Exception e){
				e.printStackTrace();
				 OutRunTimeLog.showLog(e.getMessage());
				//System.out.println(e.getMessage());
			}finally{
				return count;
			}	
		}
		
		@SuppressWarnings("finally")
		public int setFQEPQA30000Param(PreparedStatement ips, File sourceFile, String monitorId,int fDatePOS,String divCode) throws Exception{
			int count=0;		
			String pn = null;	//型号
			String wo = null;	//Lot号
			String testItem = null;	//测试项目
			String custCode = null;	//客户代码
			String testTime = null;	//时间
			String period = null;	//周期
			String testBy = null;	//测试员工
			String testPos = null;	//测量位置
			String found = null;	//实测值
			String standard = null;	//标准值
			String positive = null;	//正公差
			String negative = null;	//负公差
			String difference = null;	//差值
			String ultra = null;	//超差值
			String machNO = monitorId;	//机器编号
			
			String[] patternStr;	//pattern Array
			String encodeType=FileOptionUtil.getTXTEncodeType(sourceFile);
			Pattern pattern01 = Pattern.compile("[_.]");	//file name pattern
			Pattern pattern02 = Pattern.compile("测量项目：|客户代码：|时间：|型号：|周期：|员工：");	//operating info pattern
			Pattern pattern03 = Pattern.compile("\t{1}");
			wo = pattern01.split(sourceFile.getName().substring(sourceFile.getName().indexOf("_")+1, sourceFile.getName().length()))[1];
			try{
				String[] lines = FileOptionUtil.readLines(sourceFile,encodeType);
				for (String line : lines) {
					line = line + "\n";
					if(pattern02.split(line).length == 7){
						patternStr = pattern02.split(line);
						testItem = ParseUtil.getObjStr(patternStr[1]);
						custCode = ParseUtil.getObjStr(patternStr[2]);
						testTime = ParseUtil.getObjStr(patternStr[3]);
						pn = ParseUtil.getObjStr(patternStr[4]);
						period = ParseUtil.getObjStr(patternStr[5]);
						testBy = ParseUtil.getObjStr(patternStr[6]);
					}else if(testItem != null && pattern03.split(line).length%7 == 0){
						patternStr = pattern03.split(line);
						for(int i=0;i<patternStr.length/7;i++){
							testPos = ParseUtil.getObjStr(patternStr[7*i+0]).replace("\"", "");
							found = ParseUtil.getObjStr(patternStr[7*i+1]);
							standard = ParseUtil.getObjStr(patternStr[7*i+2]);
							positive = ParseUtil.getObjStr(patternStr[7*i+3]);
							negative = ParseUtil.getObjStr(patternStr[7*i+4]);
							difference = ParseUtil.getObjStr(patternStr[7*i+5]);
							ultra = ParseUtil.getObjStr(patternStr[7*i+6]);
							
							int ind = 1;
							ips.setDate(ind++, new java.sql.Date(new SimpleDateFormat("yyyy-MM-dd HH:mm").parse(testTime).getTime()));
							ips.setString(ind++, testItem);
							ips.setString(ind++, pn);
							ips.setString(ind++, wo);	
							ips.setString(ind++, custCode);
							ips.setString(ind++, testBy);
							ips.setString(ind++, testPos);
							ips.setString(ind++, found);
							ips.setString(ind++, standard);
							ips.setString(ind++, positive);
							ips.setString(ind++, negative);
							ips.setString(ind++, difference);
							ips.setString(ind++, ultra);
							ips.setString(ind++, period);
							ips.setString(ind++, machNO);
							ips.setString(ind++, ParseUtil.getTimeStampStr(new Date()));
							ips.setString(ind++, divCode);
							ips.addBatch();
							count++;
						}
					}
				}
			}catch (Exception e){
				e.printStackTrace();
				 OutRunTimeLog.showLog(e.getMessage());
				//System.out.println(e.getMessage());
			}finally{
				return ++count;
			}	
		}
		
		@SuppressWarnings("finally")
		public int setFQSFQA005103Param(PreparedStatement ips, File sourceFile,
				String monitorId, int fDatePOS, String divCode) throws IOException {
			int count=0;								
			String[] patternStr;	//pattern Array
			String encodeType=FileOptionUtil.getTXTEncodeType(sourceFile);
			List<String> common =new ArrayList<String>();
			Pattern head1 = Pattern.compile("客户名称:|客户代码:|终端客户名:|厂内型号:|流程卡号:");	//operating info pattern
			Pattern head2 = Pattern.compile("客户型号:|检查日期:|班次:|检查人:|审核人:");
			Pattern head3 = Pattern.compile("产品周期:|产品类型:|序号:|项目:|检测内容:");
			Pattern head4 = Pattern.compile("客户要求单位:|客户要求规格:|客户要求规格-\\+\\(公差\\):|客户要求规格--\\(公差\\):|MI要求单位:");
			Pattern head5 = Pattern.compile("是否符合:|问ICS:|附页说明:");
			Pattern pattern03 = Pattern.compile("\t{1}");
			String mirrorLine="";
			//IR索引
			List<Integer> IRbefIndex=new ArrayList<Integer>();
			List<Integer> IRaftIndex=new ArrayList<Integer>();
			List<Integer> allIndex=new ArrayList<Integer>();
			//PAD或空间距数据
			Map<String,String> PADIRBef=new HashMap<String,String>();
			Map<String,String> PADIRAft=new HashMap<String,String>();
			Map<String,String> HolespaBef=new HashMap<String,String>();
			Map<String,String> HolespaAft=new HashMap<String,String>();
			List<String> totalKeys=new ArrayList<String>();
			//正负公差
			String positive = null;	//正公差
			String negative = null;	//负公差
			String location=null;//测量位置
			String key=null;
			//PAD和孔间距比较值
			double HoCompare=0.038;
			//定义多少数据为一组
			int groupSize=5;
			//遍历循环下标
			int begin;
			int end;
			StringBuffer oneRecord=new StringBuffer();
			try{
				String[] lines = FileOptionUtil.readLines(sourceFile,encodeType);
				//寻找IR索引
				for(int i=0;i<lines.length;i++){
					mirrorLine=lines[i];
					if(head3.split(mirrorLine).length == 6){
						patternStr = head3.split(mirrorLine);
						allIndex.add(i);
						if(patternStr[5].trim().indexOf("IR前")>-1){
							IRbefIndex.add(i);
						}else if(patternStr[5].trim().indexOf("IR后")>-1){
							IRaftIndex.add(i);
						}
					}
				}
				//遍历数组
				for(int j=0;j<allIndex.size();j++){
					begin=allIndex.get(j);
					if(allIndex.size()==1||j+1==allIndex.size()){
						end=lines.length;
					}else{
						end=allIndex.get(j+1);
					}
					for(int k=begin;k<end;k++){
						//IR前
						lines[k]+="\n";
						if(pattern03.split(lines[k]).length%7 == 0&&IRbefIndex.contains(begin)){
							patternStr = pattern03.split(lines[k]);
							for(int i=0;i<patternStr.length/7;i++){
								positive = ParseUtil.getObjStr(patternStr[7*i+3]);
								negative = ParseUtil.getObjStr(patternStr[7*i+4]);
								location=ParseUtil.getObjStr(patternStr[7*i+0]).replace("\"", "");
								key=location+","+positive+","+negative;
								//PADIR前
								if(Double.valueOf(positive)>HoCompare&&Double.valueOf(negative)>HoCompare){
									if(PADIRBef.containsKey(key)){
										PADIRBef.put(key,PADIRBef.get(key)+","+ParseUtil.getObjStr(patternStr[7*i+1]));
									}else{
										//key测量位置,value=标准值,正公差,负公差,实测值
										PADIRBef.put(key,ParseUtil.getObjStr(patternStr[7*i+1]));
										if(!totalKeys.contains(key)){
											totalKeys.add(key);
										}
									}
								//孔间距IR前
								}else if(Double.valueOf(positive)==HoCompare&&Double.valueOf(negative)==HoCompare){
									if(HolespaBef.containsKey(key)){
										HolespaBef.put(key,HolespaBef.get(key)+","+ParseUtil.getObjStr(patternStr[7*i+1]));
									}else{
										//key测量位置,value=标准值,正公差,负公差,实测值
										HolespaBef.put(key,ParseUtil.getObjStr(patternStr[7*i+1]));
										if(!totalKeys.contains(key)){
											totalKeys.add(key);
										}
									}
								}
							}
						}else if(pattern03.split(lines[k]).length%7 == 0&&IRaftIndex.contains(begin)){  //IR后
							patternStr = pattern03.split(lines[k]);
							for(int i=0;i<patternStr.length/7;i++){
								positive = ParseUtil.getObjStr(patternStr[7*i+3]);
								negative = ParseUtil.getObjStr(patternStr[7*i+4]);
								location=ParseUtil.getObjStr(patternStr[7*i+0]).replace("\"", "");
								key=location+","+positive+","+negative;
								//PADIR后
								if(Double.valueOf(positive)>HoCompare&&Double.valueOf(negative)>HoCompare){
									if(PADIRAft.containsKey(key)){
										PADIRAft.put(key,PADIRAft.get(key)+","+ParseUtil.getObjStr(patternStr[7*i+1]));
									}else{
										//key测量位置,value=标准值,正公差,负公差,实测值
										PADIRAft.put(key,ParseUtil.getObjStr(patternStr[7*i+1]));
										if(!totalKeys.contains(key)){
											totalKeys.add(key);
										}
									}
								//孔间距IR后
								}else if(Double.valueOf(positive)==HoCompare&&Double.valueOf(negative)==HoCompare){
									if(HolespaAft.containsKey(key)){
										HolespaAft.put(key,HolespaAft.get(key)+","+ParseUtil.getObjStr(patternStr[7*i+1]));
									}else{
										//key测量位置,value=标准值,正公差,负公差,实测值
										HolespaAft.put(key,ParseUtil.getObjStr(patternStr[7*i+1]));
										if(!totalKeys.contains(key)){
											totalKeys.add(key);
										}
									}
								}
							}
						}
					}
				}
				//获取公共部分
				for (String line : lines) {
					line = line + "\n";
					if(head1.split(line).length == 6){
						patternStr = head1.split(line);
						for(int i=1;i<patternStr.length;i++){
							common.add(patternStr[i].trim());
						}
					}else if(head2.split(line).length == 6){
						patternStr = head2.split(line);
						for(int i=1;i<patternStr.length;i++){
							common.add(patternStr[i].trim());
						}
					}else if(head3.split(line).length == 6){
						patternStr = head3.split(line);
						for(int i=1;i<patternStr.length;i++){
							common.add(patternStr[i].trim());
						}
					}else if(head4.split(line).length == 6){
						patternStr = head4.split(line);
						for(int i=1;i<patternStr.length;i++){
							common.add(patternStr[i].trim());
						}
					}else if(head5.split(line).length==4){
						patternStr = head5.split(line);
						for(int i=1;i<patternStr.length;i++){
							common.add(patternStr[i].trim());
						}
					}
					if(common.size()==23){
						break;
					}
				}
				
				//拼凑SQL
				int maxLength=totalKeys.size();
				for(int i=0;i<maxLength;i++){
					String HolespaBefValue=null;
					String HolespaAftValue=null;
					String PADIRBefValue=null;
					String PADIRAftValue=null;
					String totalIndex=null;
					String [] HolespaBefValueArray=null;
					String [] HolespaAftValueArray=null;
					String [] PADIRBefValueArray=null;
					String [] PADIRAftValueArray=null;
					totalIndex=totalKeys.get(i);
					//孔间距前和孔间距后
					HolespaBefValue=getValueByKey(HolespaBef,totalIndex);
					HolespaAftValue=getValueByKey(HolespaAft,totalIndex);
					PADIRBefValue=getValueByKey(PADIRBef,totalIndex);
					PADIRAftValue=getValueByKey(PADIRAft,totalIndex);
					oneRecord.append(HolespaBefValue+"fancai"+HolespaAftValue+"fancai"+PADIRBefValue+"fancai"+PADIRAftValue);
					String [] records=oneRecord.toString().split("fancai");
					//最长的数据组数
					int maxRows=0;
					for(String str:records){
						int length=str.split(",").length;
						if(maxRows<length){
							maxRows=length;
						}
					}
					int InsertCount=(int) Math.ceil(maxRows/groupSize);
					HolespaBefValueArray=getGroupValue(HolespaBefValue,maxRows);
					HolespaAftValueArray=getGroupValue(HolespaAftValue,maxRows);
					PADIRBefValueArray=getGroupValue(PADIRBefValue,maxRows);
					PADIRAftValueArray=getGroupValue(PADIRAftValue,maxRows);
					for(int groupIndex=1;groupIndex<=InsertCount;groupIndex++){
						ips=setCommons(ips,common,fDatePOS);
						int index=common.size()-3;
						ips.setString(++index,"MI要求-规格");
						ips.setString(++index,"MI要求-+(公差)");
						ips.setString(++index,"MI要求--(公差)");
						for(int HolespaBefIndex=groupSize*groupIndex-groupSize;HolespaBefIndex<=groupSize*groupIndex-1;HolespaBefIndex++){
							ips.setString(++index,HolespaBefValueArray[HolespaBefIndex]);
						}
						for(int HolespaAftIndex=groupSize*groupIndex-groupSize;HolespaAftIndex<=groupSize*groupIndex-1;HolespaAftIndex++){
							ips.setString(++index,HolespaAftValueArray[HolespaAftIndex]);
						}
						for(int PADIRBeIndex=groupSize*groupIndex-groupSize;PADIRBeIndex<=groupSize*groupIndex-1;PADIRBeIndex++){
							ips.setString(++index,PADIRBefValueArray[PADIRBeIndex]);
						}
						for(int PADIRAftIndex=groupSize*groupIndex-groupSize;PADIRAftIndex<=groupSize*groupIndex-1;PADIRAftIndex++){
							ips.setString(++index,PADIRAftValueArray[PADIRAftIndex]);
						}
						ips.setString(++index,common.get(20));
						ips.setString(++index,common.get(21));
						ips.setString(++index,common.get(22));
						ips.setString(++index,divCode);
						ips.addBatch();
						count++;
					}
					oneRecord.setLength(0);
				}
			}catch (Exception e){
				e.printStackTrace();
				//System.out.println(e.getMessage());
			}finally{
				PADIRBef.clear();
				PADIRAft.clear();
				HolespaBef.clear();
				HolespaAft.clear();
				return ++count;
			}
		}

		
		@SuppressWarnings("finally")
		public int setFQSFQA005103AParam(PreparedStatement ips, File sourceFile,
				String monitorId, int fDatePOS, String divCode) throws IOException {

			int count=0;								
			String[] patternStr;	//pattern Array
			String encodeType=FileOptionUtil.getTXTEncodeType(sourceFile);
			List<String> common =new ArrayList<String>();
			Pattern head1 = Pattern.compile("客户代码:|厂内型号:|流程卡号:|检查日期:|班次:");	//operating info pattern
			Pattern head2 = Pattern.compile("检查人:|产品周期:|产品类型:|单位:|是否符合:");
			Pattern head3 = Pattern.compile("问ICS:|备注:|测量内容:");
			Pattern pattern03 = Pattern.compile("\t{1}");
			String mirrorLine="";
			//IR索引
			List<Integer> IRbefIndex=new ArrayList<Integer>();
			List<Integer> IRaftIndex=new ArrayList<Integer>();
			List<Integer> allIndex=new ArrayList<Integer>();
			//IR前和IR后
			Map<String,String>  IRBefore=new HashMap<String,String>();
			Map<String,String>  IRAfter=new HashMap<String,String>();
			List<String> totalKeys=new ArrayList<String>();
			//正负公差
			String positive = null;	//正公差
			String negative = null;	//负公差
			String location=null;//测量位置
			String standard=null; //标准值
			String key=null;
			//定义多少数据为一组
			int groupSize=5;
			//遍历循环下标
			int begin;
			int end;
			StringBuffer oneRecord=new StringBuffer();
			try{
				String[] lines = FileOptionUtil.readLines(sourceFile,encodeType);
				//寻找IR索引
				for(int i=0;i<lines.length;i++){
					mirrorLine=lines[i];
					if(head3.split(mirrorLine).length == 4){
						patternStr = head3.split(mirrorLine);
						allIndex.add(i);
						if(patternStr[3].trim().indexOf("IR前")>-1){
							IRbefIndex.add(i);
						}else if(patternStr[3].trim().indexOf("IR后")>-1){
							IRaftIndex.add(i);
						}
					}
				}
				//遍历数组
				for(int j=0;j<allIndex.size();j++){
					begin=allIndex.get(j);
					if(allIndex.size()==1||j+1==allIndex.size()){
						end=lines.length;
					}else{
						end=allIndex.get(j+1);
					}
					for(int k=begin;k<end;k++){
						//IR前
						lines[k]+="\n";
						if(pattern03.split(lines[k]).length%7 == 0&&IRbefIndex.contains(begin)){
							patternStr = pattern03.split(lines[k]);
							for(int i=0;i<patternStr.length/7;i++){
								positive = ParseUtil.getObjStr(patternStr[7*i+3]);
								negative = ParseUtil.getObjStr(patternStr[7*i+4]);
								standard=ParseUtil.getObjStr(patternStr[7*i+2]);
								location=ParseUtil.getObjStr(patternStr[7*i+0]).replace("\"", "");
								key=location+","+standard+","+positive+","+negative;
								if(IRBefore.containsKey(key)){
									IRBefore.put(key,IRBefore.get(key)+","+ParseUtil.getObjStr(patternStr[7*i+1]));
								}else{
									//key测量位置,value=标准值,正公差,负公差,实测值
									IRBefore.put(key,ParseUtil.getObjStr(patternStr[7*i+1]));
									if(!totalKeys.contains(key)){
										totalKeys.add(key);
									}
								}
							}
						}else if(pattern03.split(lines[k]).length%7 == 0&&IRaftIndex.contains(begin)){  //IR后
							patternStr = pattern03.split(lines[k]);
							for(int i=0;i<patternStr.length/7;i++){
								positive = ParseUtil.getObjStr(patternStr[7*i+3]);
								negative = ParseUtil.getObjStr(patternStr[7*i+4]);
								standard=ParseUtil.getObjStr(patternStr[7*i+2]);
								location=ParseUtil.getObjStr(patternStr[7*i+0]).replace("\"", "");
								key=location+","+standard+","+positive+","+negative;
								//IR后
								if(IRAfter.containsKey(key)){
									IRAfter.put(key,IRAfter.get(key)+","+ParseUtil.getObjStr(patternStr[7*i+1]));
								}else{
									//key测量位置,value=标准值,正公差,负公差,实测值
									IRAfter.put(key,ParseUtil.getObjStr(patternStr[7*i+1]));
									if(!totalKeys.contains(key)){
										totalKeys.add(key);
									}
								}
								
							}
						}
					}
				}
				//获取公共部分
				for (String line : lines) {
					line = line + "\n";
					if(head1.split(line).length == 6){
						patternStr = head1.split(line);
						for(int i=1;i<patternStr.length;i++){
							common.add(patternStr[i].trim());
						}
					}else if(head2.split(line).length == 6){
						patternStr = head2.split(line);
						for(int i=1;i<patternStr.length;i++){
							common.add(patternStr[i].trim());
						}
					}else if(head3.split(line).length == 4){
						patternStr = head3.split(line);
						for(int i=1;i<patternStr.length;i++){
							common.add(patternStr[i].trim());
						}
					}
					if(common.size()==13){
						break;
					}
				}

				//拼凑SQL
				String direction=null;  //方向
				String [] keyArray=null;
				int maxLength=totalKeys.size();
				for(int i=0;i<maxLength;i++){
					String BefValue=null;
					String AftValue=null;
					String totalIndex=null;
					String [] BefValueArray=null;
					String [] AftValueArray=null;
					totalIndex=totalKeys.get(i);
					keyArray=totalIndex.split(",");
					direction=keyArray[0].indexOf("X")>-1?"X":"Y";
					//IR前和IR后
					BefValue=getValueByKey(IRBefore,totalIndex);
					AftValue=getValueByKey(IRAfter,totalIndex);
					oneRecord.append(BefValue+"fancai"+AftValue);
					
					String [] records=oneRecord.toString().split("fancai");
					//最长的数据组数
					int maxRows=0;
					for(String str:records){
						int length=str.split(",").length;
						if(maxRows<length){
							maxRows=length;
						}
					}
					int InsertCount=(int) Math.ceil(maxRows/groupSize);
					BefValueArray=getGroupValue(BefValue,maxRows);
					AftValueArray=getGroupValue(AftValue,maxRows);
					for(int groupIndex=1;groupIndex<=InsertCount;groupIndex++){
						ips=setCommons(ips,common,fDatePOS);
						int index=common.size()-4;
						ips.setString(++index,direction);  //方向
						ips.setString(++index,keyArray[1]);  //规格
						ips.setString(++index,keyArray[2]);  //正公差
						ips.setString(++index,keyArray[3]);  //负公差
						for(int BefIndex=groupSize*groupIndex-groupSize;BefIndex<=groupSize*groupIndex-1;BefIndex++){
							ips.setString(++index,BefValueArray[BefIndex]);
						}
						for(int AftIndex=groupSize*groupIndex-groupSize;AftIndex<=groupSize*groupIndex-1;AftIndex++){
							ips.setString(++index,AftValueArray[AftIndex]);
						}
						ips.setString(++index,common.get(9));
						ips.setString(++index,common.get(10));
						ips.setString(++index,common.get(11));
						ips.setString(++index,divCode);
						ips.addBatch();
						count++;
					}
					oneRecord.setLength(0);
				}
			}catch (Exception e){
				e.printStackTrace();
				//System.out.println(e.getMessage());
			}finally{
				IRBefore.clear();
				IRAfter.clear();
				return ++count;
			}
		
		}		
		//获取组内的值
		private  String[] getGroupValue(String samples,int maxRows) {
			String [] result=new String [maxRows];
			String [] temp=null;
			for(int l=0;l<maxRows;l++){
				temp=samples.split(",",maxRows);
				if(l<temp.length){
					if(!"".equals(temp[l])){
						result[l]=temp[l];
					}else{
						result[l]="0";
					}
				}else{
					result[l]="0";
				}
			}
			return result;
		}
		
		//设置公共部分
		private PreparedStatement setCommons(PreparedStatement ips,
				List<String> common,int fDatePOS) throws SQLException, ParseException {
			int index=0;
			for(int j=0;j<common.size()-4;j++){
				if(j!=fDatePOS-1){
					ips.setString(++index,common.get(j));
				}else{
					try{
						java.sql.Date  d=new java.sql.Date(new SimpleDateFormat("yyyy-MM-dd HH24:mm").parse(common.get(j)).getTime());
						ips.setDate(++index,d);
					}catch(Exception e){
						ips.setTimestamp(++index, new java.sql.Timestamp(System.currentTimeMillis()));
					}
				}
			}
			return ips;
		}
		
		//根据Key获取map的值
		private String getValueByKey(Map<String, String> map,
				String keyParam) {
			String value="";
			Iterator<Entry<String, String>> iter = map.entrySet().iterator();
			while (iter.hasNext()) {
				Map.Entry entry = (Map.Entry) iter.next(); 
				String key = (String) entry.getKey();
				if(key.equals(keyParam)){
					value=(String) entry.getValue();
				}
			}
			return value;
		}
		
}
