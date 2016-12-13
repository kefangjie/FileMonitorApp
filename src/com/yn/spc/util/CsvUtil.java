package com.yn.spc.util;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class CsvUtil { 


	public String filename = null; 
	private BufferedReader bufferedreader = null; 
	private List list =new ArrayList(); 
	
	public CsvUtil(){ 
	
	} 

	
	public CsvUtil(File currFile) throws IOException{ 	
	      this.filename = currFile.getName(); 	    
	      bufferedreader =  new BufferedReader(new InputStreamReader(new FileInputStream(currFile)));
	      String stemp; 
	      while((stemp = bufferedreader.readLine()) != null){ 	     
	             list.add(stemp); 
	       } 
	} 

	public CsvUtil(String currFile) throws IOException{ 	
	      this.filename = currFile;	     
	      bufferedreader = new BufferedReader(new FileReader(filename)); 
	      String stemp; 
	      while((stemp = bufferedreader.readLine()) != null){ 	     
	             list.add(stemp); 
	       } 
	} 
	
	public List getList() throws IOException { 
	
	        return list; 
	} 
    

	//得到csv文件的行数 
	public int getRowNum(){ 
	
	        return list.size(); 
	} 


	//得到csv文件的列数 
	public int getColNum(){ 
	
	       if(!list.toString().equals("[]")) { 	       
	            if(list.get(0).toString().contains(",")) { //csv文件中，每列之间的是用','来分隔的 
	                    return list.get(0).toString().split(",").length; 
	            }else if(list.get(0).toString().trim().length() != 0) { 
	                return 1; 
	            }else{ 
	                return 0; 
	                  } 
	            }else{ 
	                return 0; 
	        } 
	} 



	//取得指定行的值 
	
	public String getRow(int index) { 
	
         if (this.list.size() != 0) return (String) list.get(index); 
         else                      
        	 return null; 
	} 


	//取得指定列的值 
	public String getCol(int index){ 
	
	       if (this.getColNum() == 0){ 
	                return null; 
	       } 
	       
	      StringBuffer scol = new StringBuffer(); 
	      String temp = null; 
	      int colnum = this.getColNum(); 
	      
	      if (colnum > 1){ 
	         for (Iterator it = list.iterator(); it.hasNext();) { 
	              temp = it.next().toString(); 
	              scol = scol.append(temp.split(",")[index] + ","); 
	          } 
	      }else{ 
	           for (Iterator it = list.iterator(); it.hasNext();) { 
	                temp = it.next().toString(); 
	                scol = scol.append(temp + ","); 
	            } 
	      } 
	        String str=new String(scol.toString()); 
	        str = str.substring(0, str.length() - 1); 
	        return str; 
	} 


	//取得指定行，指定列的值 
	public String getString(int row, int col) { 	
	        String temp = null; 
	        int colnum = this.getColNum(); 
	        
	        if(colnum > 1){ 
                temp = list.get(row).toString().split(",")[col]; 
                temp = temp.replaceAll("\"", "");
                temp = temp.trim();
	        }else if(colnum == 1) { 
                temp = list.get(row).toString(); 
                temp = temp.replaceAll("\"", "");
                temp = temp.trim();
	        }else{ 
	                temp = null; 
	        } 
	        return temp; 
	 } 


	public void CsvClose() throws IOException { 
	         this.bufferedreader.close(); 
	} 

	public void run(String filename) throws IOException { 
	         CsvUtil cu = new CsvUtil(filename); 
	         
	         BufferedWriter writer=null; 
	         try {   
	                 writer=new BufferedWriter(new FileWriter("d://tesst.txt",true));//true表示往文件后面写，不会覆盖原有内容   
	          } catch (IOException e) {   
	                e.printStackTrace(); 
	          } 
	  
	         for(int i=0;i<cu.getRowNum();i++){ 	
		           String SSCCTag = cu.getString(i,2);//得到第i行.第一列的数据. 
		         
		           String SiteName = cu.getString(i,19);//得到第i行.第二列的数据. 
		       
		           String StationId= cu.getString(i,20); 		

	              // System.out.println(SSCCTag+"    "+SiteName+"    "+StationId); 
	              try {   
	                   writer.write(SSCCTag+"       "+SiteName+"       "+StationId);  
	                    writer.newLine(); 
	                    writer.flush(); 
	              
	              } catch (IOException e) {   
	                 e.printStackTrace(); 
	               } 
				//  System.out.println("===SSCC Tag:"+SSCCTag); 
				// System.out.println("===Site Name:"+SiteName); 
				//System.out.println("===Station Id:"+StationId); 				
	
	         } 
	         
	         try {   
	             writer.close(); 
	             
	         } catch (IOException e) {   
	                e.printStackTrace(); 
	          } 
	         cu.CsvClose(); 
	} 

	public static void main(String[] args) throws IOException { 
		
		CsvUtil test = new CsvUtil(); 
		try { 
			File path = new File("D:\\SPC"); 
			File[] f = path.listFiles(); 
			List l = new ArrayList(); 
			
			for(int i=0;i<f.length;i++){ 
	          	if(f[i].getName().endsWith(".csv")) 
		           l.add(f[i]); 
		     } 
			Iterator it = l.iterator(); 
			while(it.hasNext()){ 
				File ff = (File)it.next(); 				
			    test.run(path.toString()+File.separator+ff.getName()); 
		    } 
		 } catch (Exception e) { 
		            
		 } 
		

	} 

}