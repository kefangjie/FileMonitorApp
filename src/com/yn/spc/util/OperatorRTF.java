package com.yn.spc.util;



import java.io.BufferedReader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.io.PrintWriter;
import java.io.UnsupportedEncodingException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;

import javax.swing.text.BadLocationException;
import javax.swing.text.DefaultStyledDocument;
import javax.swing.text.rtf.RTFEditorKit;

public class OperatorRTF {
		/**
		* 字符串转换为RTF编码
		* 2009-7-5  增加中文支持
		* @param content
		* @return
		*/
      int inext=0;//用来判断中文 编码出现 第一次出现为0 第二次出现为1 add by wde
	  public String strToRtf(String content){   
		  
		   char[] digital = "0123456789ABCDEF".toCharArray();
	        StringBuffer sb = new StringBuffer("");
	        byte[] bs = content.getBytes(); 
	        int bit;       
	        for (int i = 0; i < bs.length; i++) {
	            bit = (bs[i] & 0x0f0) >> 4;   
	        /*2009-7-5 add by wde 增加中文支持
	         *思路：通过getBytes获取的中文的assii小于0，根据rtf中文的的编码
	         * 所以只需要在中文的2个编码 第一个编码前加    第二个编码后加
	         * 加了一个变量inext 用来判断中文的assii 前一个和后一个。
	         * 这样在rtf中文的乱码就可以解决了。
	         */ 
	            if(bs[i]>0){
	              sb.append("\\'"); 
	            }else{               
	               if(inext==0){
	                sb.append("\\loch\\af2\\hich\\af2\\dbch\\f31505");
	                sb.append("\\'"); 
	                   inext=1;
	              }else{
	               sb.append("\\'");                   
	              }
	            }        
	            sb.append(digital[bit]);
	            bit = bs[i] & 0x0f;
	            sb.append(digital[bit]);
	            if(bs[i]<0&&inext==1){
	             sb.append("\\hich\\af2\\dbch\\af31505\\loch\\f2");
	            inext=0;
	            } 
	        }
		        
		 return sb.toString();       
	}
	  
	/**
	* 替换文档的可变部分 add by wde 2009-7-6
	* @param content 原来的文本
	* @param markersign 标记符号
	* @param replacecontent 替换的内容 
	* 用replacecontent替换markersign
	* @return
	*/
	public String replaceRTF(String content,String markersign,String replacecontent){
	   String rc = strToRtf(replacecontent);  
	   String target = "";
	   markersign="$"+markersign+"$";
	   target = content.replace(markersign,rc);     
	   return target;
	}
   public  void readRTF(String inputPath){
	   String sourname = inputPath+"\\"+"445454.rtf";
	   
	try {
		 File f = new File(sourname);
		 InputStreamReader read = new InputStreamReader (new FileInputStream(f),"UTF8");
		 BufferedReader reader=new BufferedReader(read);
		 String line;
		 System.out.println("OK==");
		 try {
			while ((line = reader.readLine()) != null) {
				
			     System.out.println(line);
			 }
		} catch (IOException e) {
			
			e.printStackTrace();
		}
		
	} catch (UnsupportedEncodingException e) {		
		e.printStackTrace();
	} catch (FileNotFoundException e) {		
		e.printStackTrace();
	}
   }
	
   /*替换模板 add by wde 2009-7-6
     * @param inputPath
	 * @param outPath
	 * @param data
	 * @return
     */
	public void rgModel(String inputPath, String outPath, HashMap data) {
	   // TODO Auto-generated method stub
	   /* 字节形式读取模板文件内容,将结果转为字符串 */
	   String sourname = inputPath+"\\"+"123.rtf";
	   String sourcecontent = "";
	   InputStream ins = null;
	   try{
	    ins = new FileInputStream(sourname);
	     byte[] b = new byte[1638400];//提高对于RTF文件的读取速度，特别是对于1M以上的文件     
	     if(ins == null){
	        System.out.println("源模板文件不存在");
	      }
	     //InputStreamReader read = ins.read();
         int bytesRead = 0;
         while (true) {
             //bytesRead = ins.read(b, 0, 1024); // return final read bytes counts
            bytesRead = ins.read(b,0,1638400);
             if(bytesRead == -1) {// end of InputStream
                  System.out.println("读取模板文件结束");
                  break;
             }
             sourcecontent += new String(b, 0, bytesRead); // convert to string using bytes
             sourcecontent = sourcecontent.replaceAll("par \\hich\\af0\\dbch\\af13\\loch\\f0", "");
             System.out.println(sourcecontent);
          }
	   }catch(Exception e){
	      e.printStackTrace();
	   }
	   
	   /* 修改变化部分 */
	   String targetcontent = "";
	   String oldText="";
	   Object newValue;
	   /* 结果输出保存到文件 */
	   try { 
		    Iterator keys = data.keySet().iterator();
		    int keysfirst=0;
		     while (keys.hasNext()){
	                oldText = (String) keys.next();
	                newValue = data.get(oldText);
	                String newText = (String) newValue; 
	                inext=0;//add by wde 改为初始状态
				     if(keysfirst==0){    
				         targetcontent = replaceRTF(sourcecontent,oldText,newText);
				         keysfirst=1;
				     }else{
				         targetcontent = replaceRTF(targetcontent,oldText,newText); 
				         keysfirst=1;
				     }  
				     System.out.println(targetcontent);
		       }  
	         
		    FileWriter fw = new FileWriter(outPath,true);
	        PrintWriter out = new PrintWriter(fw);
	        if(targetcontent.equals("")||targetcontent==""){
	           out.println(sourcecontent);
	        }else{
	           out.println(targetcontent);
	        }
	        out.close();
	        fw.close();
	        System.out.println(outPath+" 生成文件成功");
	   } catch (IOException e) {
	    // TODO Auto-generated catch block
	    e.printStackTrace();
	   }
	}

	public static void main(String[] args) throws IOException {
	   // TODO Auto-generated method stub
	   SimpleDateFormat sdf=new java.text.SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	   Date current=new Date();
	        String targetname = sdf.format(current).substring(0,4) + "年";
	        targetname += sdf.format(current).substring(5,7) + "月";
	        targetname += sdf.format(current).substring(8,10) + "日";
	        targetname += sdf.format(current).substring(11,13) + "时";                
	        targetname += sdf.format(current).substring(14,16) + "分";
	        targetname += sdf.format(current).substring(17,19) + "秒";
	        targetname+=".rtf";
	   OperatorRTF oRTF = new OperatorRTF();
	
	       //*****************************************
	   //利用HashMap读取数据库中的数据
	        HashMap map = new HashMap();
	        map.put("timetop","张三");
	        map.put("info","0155");
	        map.put("idea","公元前2000年");
	        map.put("advice","13");
	        map.put("infosend","168");     
	}
	
	/**
	 * @param inputPath
	 * @param outPath
	 * @throws IOException
	 */
	/**
	 * @param inputPath
	 * @param outPath
	 * @throws IOException
	 */
	/**
	 * @param inputPath
	 * @param outPath
	 * @throws IOException
	 */
	public void readRTFEditorKit(String inputPath,String outPath) throws IOException{
		String bodyText = null;
        DefaultStyledDocument styledDoc = new DefaultStyledDocument();   
        try {
        	InputStream is = new FileInputStream(new File(inputPath));
        	InputStreamReader inputStreamReader = new InputStreamReader(is,"utf-8");
            new RTFEditorKit().read(inputStreamReader, styledDoc, 0); //ISO8859_1 
            byte[] b2=styledDoc.getText(0, styledDoc.getLength()).getBytes("ISO8859_1");
            bodyText = new String(b2,"GBK");    //提取文本
            is.close();
            inputStreamReader.close();
        } catch (IOException e) {            
            System.out.println("不能从RTF中摘录文本!");
        } catch (BadLocationException e) {
        	System.out.println("不能从RTF中摘录文本!");            
        }
       try { 
            FileOutputStream fos = new FileOutputStream(outPath); 
            OutputStreamWriter osw = new OutputStreamWriter(fos, "GBK"); 
            osw.write(bodyText); 
            osw.flush(); 
            fos.close();
            osw.close();
        } catch (Exception e) { 
            e.printStackTrace(); 
        }
	}
	  String text;      
	  DefaultStyledDocument styledDoc ;     
	  RTFEditorKit rtf;  
	  
	public void readRtf(File in) {         
		rtf=new RTFEditorKit();         
		styledDoc =new DefaultStyledDocument();          
		try {    
			InputStreamReader reader = new InputStreamReader (new FileInputStream(in));
			rtf.read(reader, styledDoc , 0);       //ISO8859_1 UTF-8 GBK gb2312    
			text = new String(styledDoc.getText(0, styledDoc.getLength()).getBytes("ISO8859_1"));
			//String(styledDoc .getText(0, styledDoc .getLength()));  
			System.out.println(text);          
		} catch (FileNotFoundException e) {             
			// TODO Auto-generated catch block             
			e.printStackTrace();          
		} catch (IOException e) {           
			// TODO Auto-generated catch block             
			e.printStackTrace();          
		} catch (BadLocationException e) {  
			e.printStackTrace();         
		}        
	}     
	
	public void writeRtf(File out) {         
		try {              
			 rtf.write(new FileOutputStream(out), styledDoc, 0, styledDoc.getLength());         
		} catch (FileNotFoundException e) {          
			// TODO Auto-generated catch block             
			e.printStackTrace();         
		} catch (IOException e) {            
			// TODO Auto-generated catch block             
			e.printStackTrace();          
		} catch (BadLocationException e) {             
			// TODO Auto-generated catch block               
			e.printStackTrace();        
		}     
	}

}