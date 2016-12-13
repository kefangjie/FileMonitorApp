package com.yn.spc.servers.dao;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Random;
import java.util.Timer;
import java.util.TimerTask;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.TimeUnit;
import java.util.regex.Pattern;

import javax.swing.text.BadLocationException;

import net.contentobjects.jnotify.JNotify;
import net.contentobjects.jnotify.JNotifyException;
import net.contentobjects.jnotify.JNotifyListener;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.hibernate.HibernateException;
import org.hibernate.Query;
import org.hibernate.SQLQuery;
import org.hibernate.Session;

import com.yn.spc.util.*;


import com.yn.spc.common.Keys;
import com.yn.spc.threadQueue.Basket;


public class FileMonitorDAO {
	
	
	public FileMonitorDAO() throws FileNotFoundException{
		this.log=Log.getLoger();
	}
	
	private String monitorType;
	
	String[][] monitorArray;
	
	private Logger log;	

	
	private String backupFilePath = "";
	
	//监听类型
	private int mask = JNotify.FILE_CREATED; 
	
	private int watchIdPos = 0;
	
	private int monitorIdPos = 1;
	
	private int monitorPathPos = 2;
	
	private int updatedPathPos = 3;
	
	private int monitorArrayLen = 4;
	
	private  Map<String,String> SPCTable = new HashMap<String,String>();
	
	private  Map<String,String> SPCTableCols = new HashMap<String,String>();
	
	private CopyOnWriteArrayList<String> files=new CopyOnWriteArrayList<String>(); 
	
	private ConcurrentHashMap<String,String> monitorInfo=new ConcurrentHashMap<String,String>();
	
	private SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
	
	private String bgInfo="";
	
	private Timer lazyDelay  = new Timer(true);
	
	private Timer updateTimer  =null;
	
	private Timer logTimer=new Timer(true);
	
	private ExecutorService pool;

	private boolean ServiceOpenFlag =false;

	
	private InterpretXLSFileDAO  XLSFileDAO;
	private InterpretTxtFileDAO  TXTFileDAO;
	private InterpretCVSFileDAO  CVSFileDAO;
	private InterpretCLFFileDAO  CLFFileDAO;
	
	public void init(String monitorType){
		if("upload".equals(monitorType)){
			String divCode = PropertiesUtil.getKeyValue("Monitor.DivCode");
			if(SPCTable.isEmpty()){
				SPCFormHeadList(divCode);
			}
			if(SPCTableCols.isEmpty()){
				SPCFormDetail(divCode); 
			}
			initFileOption();
		}
		//设置线程数
		String ThreadCount = PropertiesUtil.getKeyValue("Monitor.ThreadCount");
		if (ThreadCount !=null && !"".equals(ThreadCount)){
		   pool=Executors.newFixedThreadPool(Integer.parseInt(ThreadCount));
		}else{
			pool=Executors.newFixedThreadPool(48);
		}
		showBgInfo();
		
	}
	
	class SpcNotifyListener implements JNotifyListener{
		//定义监听器
		public void fileCreated(int wd, String rootPath, String name){
			final String sourceFilePath = rootPath + File.separator + name;
			log.warn("文件:"+sourceFilePath+"发生改变");
			if(FileOptionUtil.isMonitorFiles(sourceFilePath)){
				for(int i=0; i<monitorArray.length; i++){
					if(Integer.toString(wd).equals(monitorArray[i][watchIdPos])){    //触发事件是该监控目标 &&!files.contains(sourceFilePath)
						try {
							log.warn("实际监测文件:"+sourceFilePath);
							//files.add(sourceFilePath);
							
							//pool.execute(new modifyRunnable(sourceFilePath,monitorArray[i][monitorIdPos]));
							Basket  basket = new Basket(sourceFilePath,monitorArray[i][monitorIdPos]);
							FoundMonitorFile producer = new FoundMonitorFile(monitorType, basket);
						       
						    ReadMonitorFile consumer =  new ReadMonitorFile(monitorType, basket);
						    pool.submit(producer);						     
						    pool.submit(consumer);
						    ServiceOpenFlag =true;
						    
						} catch (Exception e) {
							SendError(e.getMessage());
							log.error(e.getMessage());
							e.printStackTrace();
						}
					}
				}
			}
		
		}
		
		public void fileRenamed(int wd, String rootPath, String oldName, String newName){}
		
		public void fileDeleted(int wd, String rootPath, String name){}
		
		public void fileModified(int wd, String rootPath, String name){}
	}
	
	public void startMonitor(String monitorType,String path,String backup) throws IOException, SQLException{
		
		stopMonitor();
		init(monitorType);
		this.monitorType = monitorType;
		if("copy".equals(monitorType)){
			backupFilePath=backup;
			File monitorFile = new File(path);	//监控信息保存文件
			String encodeType=FileOptionUtil.getTXTEncodeType(monitorFile);
			String[] monitorLines = FileOptionUtil.readLines(monitorFile,encodeType);	//读取监控信息文件
			Pattern pattern = Pattern.compile("\t+");	//监控目录与目标目录用若干个制表符分隔
			monitorArray = new String[monitorLines.length][monitorArrayLen];
			for(int i=0; i<monitorLines.length; i++){	//读取监控信息
				String[] monitorInfo = pattern.split(monitorLines[i]);
				String monitorId = ParseUtil.getTrimStr(monitorInfo[0]);	//监控编号
				String monitorPath = ParseUtil.getTrimStr(monitorInfo[1]);	//监控目录
				if(new File(monitorPath).exists()){
					String watchId = Integer.toString(JNotify.addWatch(monitorPath, mask, true, new SpcNotifyListener()));	//开始监控
					monitorArray[i][watchIdPos] = watchId;
					monitorArray[i][monitorIdPos] = monitorId;
					monitorArray[i][monitorPathPos] = monitorPath;
					monitorArray[i][updatedPathPos] = "";	
					SendError("监测目录:"+monitorPath+"\n");
					log.info("监测目录:"+monitorPath);
				}else{
					monitorArray[i][watchIdPos] = "noExist";
					monitorArray[i][monitorIdPos] = monitorId;
					monitorArray[i][monitorPathPos] = monitorPath;
					monitorArray[i][updatedPathPos] = "";	
					bgInfo=bgInfo+monitorPath+" not found! Please check the Monitor Path.txt!\n";
					//System.out.println(monitorPath+" not found! Please check the Monitor Path.txt! ");
					SendError(bgInfo);
					log.info(bgInfo);
				}
			}
			SynchronousMoitorInfo();
		}else if("upload".equals(monitorType)){	//上传监控
			monitorArray = new String[1][monitorArrayLen];
			backupFilePath=backup;
			String monitorPath = backup;
			String watchId = Integer.toString(JNotify.addWatch(monitorPath, mask, true, new SpcNotifyListener()));	//开始监控
			monitorArray[0][watchIdPos] = watchId;
			monitorArray[0][monitorIdPos] = "";
			monitorArray[0][monitorPathPos] = monitorPath;
			monitorArray[0][updatedPathPos] = "";
		}
		bgInfo=bgInfo+"Start Monitor successful!\n";
		SendError(bgInfo);
		log.info(bgInfo);
	
	}
	
	public void stopMonitor() throws JNotifyException, SQLException{
		if(monitorArray!=null){
			for(int i=0; i<monitorArray.length; i++){
				if(!"noExist".equals(monitorArray[i][watchIdPos])){
					JNotify.removeWatch(Integer.parseInt(monitorArray[i][watchIdPos]));
				}				
			}

			SPCTable.clear();
			SPCTableCols.clear();
			monitorArray = null;
			//files.clear();
			monitorInfo.clear();
			updateTimer.cancel();
			updateTimer=null;
		}	
		
		SendError("stop Monitor......\n");
		log.info("stop Monitor......\n");	
		SendError("stop Monitor successful!\n");
		log.info("stop Monitor successful!\n");	
		
		if(ServiceOpenFlag){
			   pool.shutdown();
		}		
	    ServiceOpenFlag =false;
	}
	

	public void initFileOption(){
		
	    XLSFileDAO = new InterpretXLSFileDAO();
		TXTFileDAO = new InterpretTxtFileDAO();
	    CVSFileDAO = new InterpretCVSFileDAO();
		CLFFileDAO = new InterpretCLFFileDAO();
	}
	
	
   public class FoundMonitorFile implements Runnable {
        private String instance;
        private Basket basket;

        public FoundMonitorFile(String instance, Basket basket) {
        	super();
            this.instance = instance;
            this.basket = basket;
        }

        public void run() {
            try {
                //while (true) { 
                	SendError(instance+"producer -bg-> " + basket.getSourceFilePath()+"\n");
                    basket.produce();                   
                    // 休眠30ms                    
                	SendError(instance+"producer -end-> " + basket.getSourceFilePath()+"\n");
					log.info("producer --> " + basket.getSourceFilePath()+"\n");
					 if("upload".equals(instance)){
                         Thread.sleep(10);
                   // }
               }
            } catch (InterruptedException ex) {
            	SendError(instance+"FoundMonitorFile 异常" + basket.getSourceFilePath()+"\n");
				log.info("FoundMonitorFile 异常" + basket.getSourceFilePath()+"\n");
            }
        }
    }
   

   // 定义苹果消费者
   class ReadMonitorFile implements Runnable {
       private String instance;
       private Basket basket;
	   private String sourceFilePath;
	   private String pathValue;
	   private String filename =null;

       public ReadMonitorFile(String instance, Basket basket) {
    	   super();
           this.instance = instance;
           this.basket = basket;
           
       }
       
       @Override
       public void run() {
           try {
        	
              while (!basket.basketQueue.isEmpty()) {             	 
                   filename = basket.consume();     
                 	SendError("Copy:"+sourceFilePath + " -bg-> " + filename+"\n");
				    log.info("Copy:"+sourceFilePath + " -end-> " + filename+"\n");
                   if(filename !=null){
                	   if("copy".equals(instance)){
	                	   sourceFilePath = filename.split("=#=")[0];
	                	   pathValue = filename.split("=#=")[1];   
                	   }else{
                		   sourceFilePath = filename.split("=#=")[0];
                	   }             	  
                	
                	   File sourceFile = new File(sourceFilePath);	//已更新文件
                		if("copy".equals(instance)){	//备份监控
        					String sToday=sdf.format(new Date());
        					String targetDirPath = backupFilePath+ File.separator +pathValue;//+ File.separator+sToday;	//目标路径
        					String targetFilePath = targetDirPath + File.separator;
        					Random random=new Random();
        					String filemark=sToday+Math.abs(random.nextInt())+"_";   //目标文件名前加上系统时间和产生的一个随机数
        					if(sourceFile.getName().toUpperCase().endsWith(".RTF")){            //将rtf转换为txt
        						targetFilePath+=filemark+sourceFile.getName().replace(".rtf", ".txt");
        					}else{
        						targetFilePath+=filemark+sourceFile.getName();
        					}
        					File targetDir = new File(targetDirPath);
        					if(!targetDir.exists()){
        						File targetDirParent=new File(targetDir.getParent());
        						if(!targetDirParent.exists()){
        							targetDirParent.mkdirs();
        						}
        						targetDir.mkdir(); 
        					}
        					File targetFile = new File(targetFilePath);	//目标文件
        					try {
        						try {
        							FileOptionUtil.copySingleFile(sourceFile, targetFile);
        							lazyDelay.schedule(new TimerTask(){
        								@Override
        								public void run() {
        									files.remove(sourceFilePath);
        								}},10);
        						} catch (InterruptedException e) {
        							SendError(e.getMessage());
        							e.printStackTrace();
        						} catch (IOException e) {
        							SendError(e.getMessage());
        							e.printStackTrace();
        						}
        					} catch (BadLocationException e) {
        						SendError(e.getMessage());
        						e.printStackTrace();
        					}finally{
        					//拷贝文件
	        					bgInfo=bgInfo+sourceFilePath + " --> " + targetFilePath+"\n";
	        					
	        					SendError("Copy:"+sourceFilePath + " -end-> " + targetFilePath+"\n");
	        					log.info("Copy:"+sourceFilePath + " -end-> " + targetFilePath+"\n");
        					}
        				}else if("upload".equals(instance)){	//上传监控
        					
        					bgInfo=bgInfo+"Begined Upload File:" + sourceFile+"\n";
        					SendError("Begined Upload File:" + sourceFile+"\n");
        					log.info("Begined Upload File:" + sourceFile+"\n");
        					try {
        						uploadFile(sourceFile,sourceFilePath);
        					} catch (Exception e) {
        						e.printStackTrace();
        					}
        					Thread.sleep(20);
        				}               	   
                	   
                   }                       
              }
           } catch (InterruptedException ex) {
           	    SendError(instance+"ReadMonitorFile 异常" + sourceFilePath+"\n");
				log.info("ReadMonitorFile 异常" + sourceFilePath+"\n");
           }
       }
   }
    
	/*
	 * 文件监控线程类
	 */
	public class modifyRunnable implements Runnable {
		
		private String sourceFilePath;
		private String pathValue;
		
		public modifyRunnable(String sourceFilePath, String pathValue) {
			super();
			this.sourceFilePath = sourceFilePath;
			this.pathValue = pathValue;
		}

		@Override
		public void run() {
				File sourceFile = new File(sourceFilePath);	//已更新文件
				if(sourceFile.isFile()){
					if("copy".equals(monitorType)){	//备份监控
						String sToday=sdf.format(new Date());
						String targetDirPath = backupFilePath + File.separator + pathValue;//+ File.separator+sToday;	//目标路径
						String targetFilePath = targetDirPath + File.separator;
						Random random=new Random();
						String filemark=sToday+Math.abs(random.nextInt())+"_";   //目标文件名前加上系统时间和产生的一个随机数
						if(sourceFile.getName().toUpperCase().endsWith(".RTF")){            //将rtf转换为txt
							targetFilePath+=filemark+sourceFile.getName().replace(".rtf", ".txt");
						}else{
							targetFilePath+=filemark+sourceFile.getName();
						}
						File targetDir = new File(targetDirPath);
						if(!targetDir.exists()){
							File targetDirParent=new File(targetDir.getParent());
							if(!targetDirParent.exists()){
								targetDirParent.mkdirs();
							}
							targetDir.mkdir(); 
						}
						File targetFile = new File(targetFilePath);	//目标文件
						try {
							try {
								FileOptionUtil.copySingleFile(sourceFile, targetFile);
								
								lazyDelay.schedule(new TimerTask(){
									@Override
									public void run() {
										files.remove(sourceFilePath);
									}},100);
							} catch (InterruptedException e) {
								SendError(e.getMessage());
								e.printStackTrace();
							} catch (IOException e) {
								SendError(e.getMessage());
								e.printStackTrace();
							}
						} catch (BadLocationException e) {
							SendError(e.getMessage());
							e.printStackTrace();
						}	//拷贝文件
						//bgInfo=bgInfo+sourceFilePath + " --> " + targetFilePath+"\n";
						//System.out.println(sourceFilePath + " --> " + targetFilePath);
						SendError(sourceFilePath + " --> " + targetFilePath+"\n");
						log.info(sourceFilePath + " --> " + targetFilePath+"\n");
					}else if("upload".equals(monitorType)){	//上传监控
						//System.out.println("Begined Upload File:" + sourceFile);
						//bgInfo=bgInfo+"Begined Upload File:" + sourceFile+"\n";
						SendError("Begined Upload File:" + sourceFile+"\n");
						log.info("Begined Upload File:" + sourceFile+"\n");
						try {
							uploadFile(sourceFile,sourceFilePath);
							
							lazyDelay.schedule(new TimerTask(){
								@Override
								public void run() {
									files.remove(sourceFilePath);
								}},100);
								
						} catch (Exception e) {
							e.printStackTrace();
						}
					}
				}else{
					SendError("不是文件:" + sourceFile+"\n");
					log.info("不是文件:" + sourceFile+"\n");
				}
			}
	}
	
		
	private void uploadFile(File sourceFile,String sourceFilePath) throws Exception{
		String[] formArray;			
		String formID = null;		
		String formTable = null;			
		String formName = null;
		String formCols;	
		String monitorId = sourceFile.getParent().replace(backupFilePath + File.separator, "");
		//monitorId=monitorId.substring(0, monitorId.length()-9).replace('/', '-');
		String divCode = PropertiesUtil.getKeyValue("Monitor.DivCode");
		if (sourceFile.isFile()){
		   int fDatePOS = 0;
		  
		   String fileEXT=sourceFile.getName().substring(sourceFile.getName().lastIndexOf(".")+1).toUpperCase();
		   if(SPCTable.containsKey(monitorId)){
				formArray = SPCTable.get(monitorId).split("=jk=");				
				formID     = formArray[0];
				formTable  = formArray[1];
				formName   = formArray[2];	
			}
			if(formID !=null && !"".equals(formID)){
				formCols = SPCTableCols.get(formID);
				for(String Field : formCols.split(",")){
					fDatePOS = fDatePOS +1;
					if(Field.toUpperCase().equals("OPTDATE"))
					   break;	
				}
				if(fileEXT.equals("TXT") || fileEXT.equals("STA")  || fileEXT.equals("EXP")){
					
					saveSPCDataFrom(monitorId, sourceFile,formID,formTable,formName, formCols, fDatePOS,divCode,"txt");
				
				}else if(fileEXT.equals("XLS") || fileEXT.equals("XLSX")){
					
					saveSPCDataFrom(monitorId, sourceFile,formID,formTable,formName, formCols, fDatePOS,divCode,"xls");
				
				}else if(fileEXT.equals("CSV")){
					
					saveSPCDataFrom(monitorId, sourceFile,formID,formTable,formName, formCols, fDatePOS,divCode,fileEXT);
				
				}else if(fileEXT.equals("CLF")){
					
					saveSPCDataFrom(monitorId, sourceFile,formID,formTable,formName, formCols, fDatePOS,divCode,fileEXT);
				
				}
				BackUpFile(monitorId, sourceFile);
			}else{
				
				SendError(monitorId+"不在SPC自动监控范围内,请确认配置是否正确?");
				excuteSQL(monitorId,"",-1,sourceFile.getName(),"非SPC自动监控项目","-1",divCode);
			}
			//files.remove(sourceFilePath);
			
		
			bgInfo=bgInfo+"Finished Upload File:" + sourceFile+"\n";
			SendError("Finished Upload File:" + sourceFile+"\n");
			log.info("Finished Upload File:" + sourceFile+"\n");
			
		}
	}
	
	private void BackUpFile(String monitorId,File sourceFile){
		String sToday=ParseUtil.getDateStr();
		String targetDirPath  = PropertiesUtil.getKeyValue("BackupDirectory");
		if(targetDirPath==null || "".equals(targetDirPath)){
			targetDirPath = Keys.backupPath + File.separator;
		}else{

			targetDirPath = targetDirPath + File.separator;
		}
		
		targetDirPath +=  monitorId+ File.separator+sToday;	//目标路径
		String targetFilePath = targetDirPath;
		File targetDir = new File(targetDirPath);
		if(!targetDir.exists()){
			File targetDirParent=new File(targetDir.getParent());
			if(!targetDirParent.exists()){
				targetDirParent.mkdirs();
			}			
			targetDir.mkdir(); 
		}
		targetFilePath = targetDirPath + File.separator+sourceFile.getName();
		File targetFile = new File(targetFilePath);	//目标文件
		try {
			try {
				FileOptionUtil.copySingleFile(sourceFile, targetFile);
		
			} catch (InterruptedException e) {
				SendError(e.getMessage());
				e.printStackTrace();
			} catch (IOException e) {
				SendError(e.getMessage());
				e.printStackTrace();
			}
		} catch (BadLocationException e) {
			SendError(e.getMessage());
			e.printStackTrace();
		}	//备份文件	
		sourceFile.delete();
		log.info(" --> " + targetFilePath+"\n");
		
	}
	
	private void SPCFormHeadList(String divCode){			
		Session session = DBOptionUtil.getSession();		
		StringBuffer hql = new StringBuffer();
		hql.append(" SELECT  replace(PARAMCODE,'/','-')PARAMCODE, (PKEY || '=jk=' || TABLENAME || '=jk=' || PARAMNAME)Pkey ");
		hql.append(" from  MTG_ProdParamHead ");
		hql.append(" where  (REMARK like '%SPC%' OR REMARK like '%MRB%' OR REMARK like '%LDM%' )");
		hql.append("       and div_code in ('"+divCode+"','MTG') ");
		hql.append(" order by PARAMCODE ");
		Query  qty=session.createSQLQuery(hql.toString());
		List<Object[]> list = (List<Object[]>)qty.list(); 
		if (list != null) {			
			SPCTable.clear();
		   for (Object[] array: list) { 			 
			   SPCTable.put(array[0].toString(), array[1].toString());
		   }
	    }
		DBOptionUtil.closeSession();
	}
	
	private void SPCFormDetail(String divCode){
		Session session = DBOptionUtil.getSession();		
		StringBuffer hql = new StringBuffer();
		hql.append(" SELECT  im.PARAMH_PTR,im.FIELDNAME  from    MTG_ProdParamItems im ");
		hql.append(" where ACTION_FLAG ='1' and exists (select 1 from  MTG_ProdParamHead hd ");
		hql.append(" where im.paramh_ptr = hd.pkey and  (hd.REMARK like '%SPC%'  OR REMARK like '%MRB%' OR REMARK like '%LDM%' )  ");
		hql.append("   and hd.div_code  in ('"+divCode+"','MTG') )");
		hql.append(" order by  im.PARAMH_PTR,im.seq ");
		Query  qty=session.createSQLQuery(hql.toString());
		List<Object[]> list = (List<Object[]>)qty.list(); 
		if (list != null) {			
			SPCTableCols.clear();
		   for (Object[] array: list) { 
				if(SPCTableCols.get(array[0].toString())!="" && 
				   SPCTableCols.get(array[0].toString())!=null){					
				   SPCTableCols.put(array[0].toString(),SPCTableCols.get(array[0].toString())+array[1].toString()+",");				
				}else{
					SPCTableCols.put(array[0].toString(), array[1].toString()+",");
				}			  
		   }
	    }
		DBOptionUtil.closeSession();
	}

	private String insertSQL(String formTable, String formCols, String formID, int fDatePOS) throws Exception{
		StringBuffer iSql = new StringBuffer();
		iSql.append("INSERT INTO " +formTable+"(PKEY, PARAMH_PTR,"+formCols);
		if(fDatePOS == 0){
			iSql.append("OPTDATE,DIV_CODE,LAST_TX_DT,LAST_TX_USERNAME) VALUES(");	
		}else{
			iSql.append("DIV_CODE,LAST_TX_DT,LAST_TX_USERNAME) VALUES(");	
		}		
		iSql.append(formTable+"_SEQ.NEXTVAL,"+formID+",:"+formCols.replace(",", ",:"));
		if(fDatePOS == 0){
			iSql.append("OPTDATE,:DIV_CODE,sysdate,'SysAuto')");	
		}else{			
			iSql.append("DIV_CODE,sysdate,'SysAuto')");	
		}
		return iSql.toString();
	}
	
	private void saveSPCDataFrom(String monitorId, File sourceFile,String formID,
			       String formTable,String formName,String formCols,int fDatePOS, String divCode, String fileType) throws Exception{
		PreparedStatement ips=null;
		String Massege="" ,CurrStatus="";
		Connection conn=DBOptionUtil.getThreadLocalConnection();
		int count = 0;
		try{
			String insertSql = insertSQL(formTable, formCols, formID, fDatePOS);
			ips= conn.prepareStatement(insertSql);			
			//System.out.println(Thread.currentThread()+""+"PreparedStatement is "+ips);
		
			if("xls".equals(fileType)){			
				count=setXLSRealParam(ips,monitorId,sourceFile,fDatePOS,divCode);
			}else if("txt".equals(fileType)){
				count=setTXTRealParam(ips,monitorId,sourceFile,fDatePOS,divCode);
			}else if("CSV".equals(fileType)){
				count=CSVFileParsing(ips,monitorId,sourceFile,fDatePOS,divCode);
			}else if("CLF".equals(fileType)){
				count=CLFDataConverter(ips,monitorId,sourceFile,fDatePOS,divCode);
			}
			if(count>=1){	
				if("F-QE-ROU-D001".equals(monitorId) && (count %5000 !=0)){
				   ips.executeBatch();
				}else{
				   ips.executeBatch();
				}
				Massege = "成功上传数据";				
				CurrStatus = "Succeed";					
			}else{
				CurrStatus = "Fail";
				Massege = "文件格式错误，无法读取数据，数据上传失败";
			}
		}catch (Exception e){
			e.printStackTrace();
			SendError(e.getMessage());
			Massege = e.getMessage().substring(0, (e.getMessage().length()>3000?3000:e.getMessage().length()));
			CurrStatus = "Fail";			
		}finally{
			if(ips !=null){
				ips.close();
				ips=null;
			}
			excuteSQL(monitorId,formName,count,sourceFile.getName(),Massege,CurrStatus,divCode);
			DBOptionUtil.closeThreadLocalConnection();
		} 	
	}
	
	@SuppressWarnings("finally")
	private int CLFDataConverter(PreparedStatement ips, String monitorId,File sourceFile,int fDatePOS,String divCode) throws Exception{
		int count=0;		
		try{
			while(true){
				if(sourceFile.renameTo(sourceFile)){
					double start  = System.currentTimeMillis() ; 
					if("F-QE-ROU-D001".equals(monitorId)){ 
						count=CLFFileDAO.ROU_CLFParam(ips, sourceFile, monitorId,fDatePOS,divCode);
					}else{	                 
						count=CLFFileDAO.CommonCLFParam(ips, sourceFile, monitorId,fDatePOS,divCode);
					}
					
					double end = System.currentTimeMillis() ;   
					SendError("CLF耗时: " + (end - start)+"秒\n");					 
					 break;
				}else{
					TimeUnit.MILLISECONDS.sleep(1000);
				}
			}

		}catch (Exception e){
			e.printStackTrace();			
		}finally{			
			return count;
		}		
	}

		
	@SuppressWarnings("finally")
	private int CSVFileParsing(PreparedStatement ips, String monitorId,File sourceFile,int fDatePOS,String divCode) throws Exception{
		int count=0;	
		CsvUtil csv =null;
		try{
			while(true){
				if(sourceFile.renameTo(sourceFile)){
					 csv = new CsvUtil(sourceFile);					
					 break;
				}else{
					TimeUnit.MILLISECONDS.sleep(1000);
				}
			}
			double start  = System.currentTimeMillis() ; 
			if("CVS-01".equals(monitorId)){ 
				//count=CommonCSVFileParsing(ips, monitorId,csv,fDatePOS,divCode);
			}else{	                 
				count=CVSFileDAO.CommonCSVFileParsing(ips, monitorId,csv,fDatePOS,divCode);	
			}
			
			double end = System.currentTimeMillis() ;           
			//System.out.println("CSV执行时间 : " + (end - start)); 
			 SendError("CSV耗时: " + (end - start)+"秒\n");
		}catch (Exception e){
			e.printStackTrace();			
		}finally{
			csv.CsvClose();
			return count;
		}		
	}
	
	@SuppressWarnings("finally")
	private int setTXTRealParam(PreparedStatement ips, String monitorId, File sourceFile,int fDatePOS,String divCode) throws Exception{
		int count=0;
		try{
			while(true){
				if(sourceFile.renameTo(sourceFile)){
					if("F-QE-PQA-30000".equals(monitorId)){         //3D 镭射检测机
						count=TXTFileDAO.setFQEPQA30000Param(ips, sourceFile, monitorId,fDatePOS,divCode);
					}else if("F-QS-FQA-005-10.3".equals(monitorId)){                 //3D 测量仪
						count=TXTFileDAO.setFQSFQA005103Param(ips, sourceFile, monitorId,fDatePOS,divCode);
					}else if("F-QS-FQA-005-10.3A".equals(monitorId)){  
						count=TXTFileDAO.setFQSFQA005103AParam(ips, sourceFile, monitorId,fDatePOS,divCode);
					}else if("F-QE-FISCHERX".equals(monitorId)){  
						count=TXTFileDAO.ImpFischerXData(ips, sourceFile, monitorId,fDatePOS,divCode);
					}else{                            //处理通用txt格式 
						count=TXTFileDAO.setCommonTXTParam(ips, sourceFile, monitorId,fDatePOS,divCode);
					}
					break;
				}else{
					TimeUnit.MILLISECONDS.sleep(1000);
				}	
			}
		}catch (Exception e){
			SendError(e.getMessage());
			e.printStackTrace();
			System.out.print(e.getMessage());
		}finally{
			return count-1;
		}					
	}	
	
	@SuppressWarnings("finally")
	private int setXLSRealParam(PreparedStatement ips, String monitorId,File sourceFile,int fDatePOS,String divCode) throws Exception{
		int count=0;
		Sheet sheet=null;
		try{
			while(true){
				if(sourceFile.renameTo(sourceFile)){
					 sheet = getHSSFSheet(sourceFile);	
					 break;
				}else{
					TimeUnit.MILLISECONDS.sleep(1000);
				}
			}
			double start  = System.currentTimeMillis() ; 
			if("F-QE-PQA-410A".equals(monitorId)){                   
				//count= XLSFileDAO.setFQEPQA410AParam2(ips, monitorId,sheet,fDatePOS,divCode);
				count=XLSFileDAO.setFQEPQA410AParam2(ips,monitorId,sheet,fDatePOS,divCode);
			}else if("F-QE-PQA-366A".equals(monitorId)){              
				count=XLSFileDAO.setFQEPQA366AParam3(ips, monitorId,sheet,fDatePOS,divCode);
			}else if("D-ME-LAM-003A".equals(monitorId)){                
				count=XLSFileDAO.setDMELAM003AParam3(ips, monitorId,sheet,fDatePOS,divCode);
			}else if("F-QE-ROU-D001".equals(monitorId)){               
				//count=XLSFileDAO.setFQEROUD001Param(ips, monitorId,sheet,fDatePOS,divCode);
				count=XLSFileDAO.setCommonExcelParam(ips, monitorId,sheet,fDatePOS,divCode);	
			}else if("F-QE-PQA-396".equals(monitorId)){                
				count=XLSFileDAO.setFQEFQA396Param(ips, monitorId,sheet,fDatePOS,divCode);
			}else if("F-QC-IDF-001".equals(monitorId)){                
				count=XLSFileDAO.ImpQCIDF01Data(ips, monitorId,sheet,fDatePOS,divCode);
			}else if("F-QC-LDR-001".equals(monitorId)){                
				count=XLSFileDAO.ImpQCLDR01Data(ips, monitorId,sheet,fDatePOS,divCode);
			}else if("F-QE-PQA-374".equals(monitorId)){                
				count=XLSFileDAO.ImpPQA374Data(ips, sourceFile.getName(), monitorId,sheet,fDatePOS,divCode);
			} else if ("F-QE-PQA-ODF-LINE".equals(monitorId)) {
		        count = XLSFileDAO.setQEPQAIPQCDesdLineWParam(monitorId, sheet, fDatePOS, divCode);
		    } else if ("F-QE-PQA-ODF-IMPD".equals(monitorId)){
		        count = XLSFileDAO.setQEPQA410CDESDParam(monitorId, sheet, fDatePOS, divCode);
		    }else{	                  // 适合通用excel格式 
				count=XLSFileDAO.setCommonExcelParam(ips, monitorId,sheet,fDatePOS,divCode);	
			}
			double end = System.currentTimeMillis() ;   
		
			//System.out.println("setXLSRealParam time is : " + (end - start)); 
			SendError("XLS耗时: " + (end - start)+"秒\n");
		}catch (Exception e){
			e.printStackTrace();			
		}finally{
			return count;
		}		
	}
	

	private  int excuteSQL(String RPT_Code,String RPT_Name,int Imp_Count,String File_Name,
			               String Reason_Failure,String Status,String DIVCODE){
		int count=0;
		Session session=DBOptionUtil.getSession();
		try {
			DBOptionUtil.beginTransaction();
            StringBuffer sql = new StringBuffer();
            sql.append(" insert into MTG_SPCMonitor_LOG(PKEY,RPT_Code,RPT_Name,Opt_Date,Imp_Direction,Imp_Count,File_Name,Reason_Failure,Status,DIV_CODE)");
            sql.append(" VALUES(MTG_SPCMonitor_LOG_SEQ.NEXTVAL,'");
            sql.append( RPT_Code+"','"+RPT_Name+"',sysdate,'TowERP',"+Imp_Count);
            sql.append( ",'"+File_Name+"','"+Reason_Failure+"','"+Status+"','"+DIVCODE+"')");
			SQLQuery  query=session.createSQLQuery(sql.toString());			
			count=query.executeUpdate();
		} catch (HibernateException e) {
			SendError(e.getMessage());
			e.printStackTrace();			
			return count;
		} 
		DBOptionUtil.commitTransaction();
		DBOptionUtil.closeSession();
		return count;
		
	}
	
	private Sheet getHSSFSheet(File sourceFile)throws Exception {
		Workbook wb=null;
		InputStream inStream=null;
		try {
			inStream = new FileInputStream(sourceFile);			
			wb=WorkbookFactory.create(inStream);
		} catch (IOException e){
			e.printStackTrace();
		}finally{
			inStream.close();
		}
		Sheet sheet = wb!=null?wb.getSheetAt(0):null;
		return sheet;
	}


	public String readConsoleInfo() throws Exception{
		String readResult="";
		SimpleDateFormat sdfd = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		if(bgInfo!=null&&!"".equals(bgInfo)){
			bgInfo="";
		}else{
			readResult=sdfd.format(new Date())+"    "+"waiting operation...";
			bgInfo="";
		}		
		return readResult;
	}
	
	private synchronized  void updateMonitorArray() throws JNotifyException{
		for(int i=0; i<monitorArray.length; i++){
			String watchId=monitorArray[i][watchIdPos];
			String monitorPath=monitorArray[i][monitorPathPos];
			String flag=monitorArray[i][updatedPathPos];
			//不存在表示重启或监控目录被删除了
			if(!new File(monitorPath).exists()&&"".equals(flag)){
				log.error("目标机器已关机:"+monitorPath+"\n");
				SendError("目标机器已关机:"+monitorPath+"\n");
				monitorArray[i][updatedPathPos]="true";
				monitorInfo.put(watchId, monitorPath);
				if(!"noExist".equals(watchId))
				   JNotify.removeWatch(Integer.parseInt(watchId));
			}			
		}
		Iterator<Entry<String, String>> iter = monitorInfo.entrySet().iterator();
		while (iter.hasNext()) {
			Map.Entry entry = (Map.Entry) iter.next(); 
			String key = (String) entry.getKey();
			String value=(String) entry.getValue();
			if(new File(value).exists()){
				String watchId = Integer.toString(JNotify.addWatch(value, mask, true, new SpcNotifyListener()));
				for(int i=0; i<monitorArray.length; i++){
					String monitorPath=monitorArray[i][monitorPathPos];
					if(monitorPath.equals(value)){
						monitorArray[i][watchIdPos]=watchId;
						monitorArray[i][updatedPathPos]="";
						log.info("目标机器已开机:"+monitorPath+"\n");
						SendError("目标机器已开机:"+monitorPath+"\n");
						monitorInfo.remove(key);
					}
				}
			}
		}
	}
	
	
	public void SynchronousMoitorInfo(){
		updateTimer=new Timer(true);
		updateTimer.schedule(new TimerTask(){
        	public void run() {
    			try {
    				updateMonitorArray();
    			} catch (Exception e) {
    				e.printStackTrace();
    			}
    		}
    	}, 1000,10000);
    }
	
	
	 public void showBgInfo(){
		 logTimer.schedule(new TimerTask(){
	        	public void run() {
	    			try {
	    				SendError(readConsoleInfo()+"\n");
	    			} catch (Exception e) {
	    				e.printStackTrace();
	    			}
	    		}
	    	}, 1000,10000);
	    }
	 
	public void  getMonitorInfo(){
		System.out.println("监控情况:");
		SendError("监控情况如下:\n");
		if(monitorArray==null||monitorArray.length==0){
			System.out.println("无监控目录");
			SendError("无监控目录\n");
			log.info("无监控目录");
		}else{
			for(int i=0;i<monitorArray.length;i++){
				System.out.println("监控ID:"+monitorArray[i][watchIdPos]+"  监控目录"+monitorArray[i][monitorPathPos]);
				SendError("监控ID:"+monitorArray[i][watchIdPos]+"  监控目录"+monitorArray[i][monitorPathPos]+"\n");
				log.info("监控ID:"+monitorArray[i][watchIdPos]+"  监控目录"+monitorArray[i][monitorPathPos]+"\n");
			}
		}
		
	}
	
    private void SendError(String mes){
 	   OutRunTimeLog.showLog(mes);	
    }
}
