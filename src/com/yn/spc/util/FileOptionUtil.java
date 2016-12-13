package com.yn.spc.util;

import java.io.BufferedInputStream;

import java.io.BufferedOutputStream;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.RandomAccessFile;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;

import java.util.ArrayList;

import java.util.List;
import java.util.concurrent.TimeUnit;
import java.util.regex.Pattern;

import javax.swing.text.BadLocationException;
import javax.swing.text.DefaultStyledDocument;
import javax.swing.text.rtf.RTFEditorKit;





public final class FileOptionUtil {
	
	//判断是否监听文件类型
	public static boolean isMonitorFiles(String sourceFilePath){
		
		String fileName = new File(sourceFilePath).getName().toUpperCase();	//已更新文件名（大写）
		
		return !fileName.startsWith("~") && (fileName.endsWith(".TXT")
				||fileName.endsWith(".XLSX")||fileName.endsWith(".XLS")
				||fileName.endsWith(".DOCX")||fileName.endsWith(".DOC")
				||fileName.endsWith(".RTF") ||fileName.endsWith(".CSV")
				||fileName.endsWith(".CLF") ||fileName.endsWith(".EXP"));
	}
	
	public static void copySingleFile(File sourceFile, File targetFile) throws BadLocationException, InterruptedException, IOException{
		String sourceFileType =sourceFile.getName().toUpperCase();
		String targetFileType=targetFile.getName().toUpperCase();
		
		if(sourceFileType.endsWith(".RTF")&&targetFileType.endsWith(".TXT")){
			rtf2txt(sourceFile.getParent()+ File.separator +sourceFile.getName(),targetFile.getParent()+ File.separator +targetFile.getName());
		}else{	
			//otherType(sourceFile,targetFile);
			nioTransferCopy(sourceFile,targetFile);
			//NIOBufferCopy(sourceFile,targetFile);
			//FileUtils.copyFile(sourceFile,targetFile);
			//customBufferBufferedStreamCopy(sourceFile,targetFile);
		}
	}
    private static void customBufferBufferedStreamCopy(File source, File target) throws IOException {  
        InputStream fis = null;  
        OutputStream fos = null;  
        try {  
         
            fis = new BufferedInputStream(new FileInputStream(source));  
            fos = new BufferedOutputStream(new FileOutputStream(target));  
            byte[] buf = new byte[4096];  
            int i;  
            while ((i = fis.read(buf)) != -1) {  
                fos.write(buf, 0, i);  
            }  
        }  
        catch (Exception e) {  
            e.printStackTrace();  
        } finally { 
        	fis.close();
        	fos.close();       
        }  
    } 
	
    private static void NIOBufferCopy(File source, File target) throws IOException {  
        FileChannel in = null;  
        FileChannel out = null;  
        FileInputStream inStream = null;  
        FileOutputStream outStream = null;  
        try {  
            inStream = new FileInputStream(source);  
            outStream = new FileOutputStream(target);  
            in = inStream.getChannel();  
            out = outStream.getChannel();  
            ByteBuffer buffer = ByteBuffer.allocate(1024*5);  
            while (in.read(buffer) != -1) {  
                buffer.flip();  
                out.write(buffer);  
                buffer.clear();  
            }  
        } catch (IOException e) {  
            e.printStackTrace();  
        } finally {  
            inStream.close();
            in.close();
            outStream.close();
            out.close();
        }  
    } 
	
	public static  int  UpLoadFileforChannel(File f1,File f2) throws IOException{		   
	    int length= (int) f1.length() ;//2097152; 
	    
	    RandomAccessFile randomFile = new RandomAccessFile(f1, "r");
	
	    //FileInputStream in=new FileInputStream(f1);
	    FileOutputStream out=new FileOutputStream(f2);
	    FileChannel inC=  randomFile.getChannel() ;//in.getChannel();
	    FileChannel outC=out.getChannel();
	    ByteBuffer b=null;
		
	    while(true){
	        if(inC.position()==inC.size()){
	            
	            inC.close();
	            outC.close();
	            out.close();
	          //  in.close(); 
	            randomFile.close();
	            return 1;
	        }
	        if((inC.size()-inC.position())<length){
	            length=(int)(inC.size()-inC.position());
	        }else
	            length=(int) f1.length() ;//2097152; 
	        b=ByteBuffer.allocateDirect(length);
	        inC.read(b);
	        b.flip();
	        outC.write(b);
	        outC.force(false);
	    }
	}
	
	public static void nioTransferCopy(File source, File target) throws IOException {  
        FileChannel in = null;  
        FileChannel out = null;  
        FileInputStream inStream = null;  
        FileOutputStream outStream = null;  
        try {  
            inStream = new FileInputStream(source);  
            outStream = new FileOutputStream(target);  
            in = inStream.getChannel();  
            out = outStream.getChannel();  
            in.transferTo(0, in.size(), out);  
        } catch (IOException e) {  
            e.printStackTrace();  
        } finally {  
          
            inStream.close();
            in.close();
            outStream.close();
            out.close();
           
        }  
    } 
    
	//其他类型的读取方式
	public static void otherType(File sourceFile, File targetFile) throws InterruptedException, IOException{
		FileInputStream input=null;
		FileOutputStream output=null;

		while(true){
			if(sourceFile.renameTo(sourceFile)){
				input = new FileInputStream(sourceFile);
				BufferedInputStream inBuff = new BufferedInputStream(input);
				output = new FileOutputStream(targetFile);
				BufferedOutputStream outBuff = new BufferedOutputStream(output);
				
				byte[] b = new byte[1024 * 10];
				int len;
				while ((len = inBuff.read(b)) != -1) {
					outBuff.write(b, 0, len);
					
				}
				outBuff.flush();		
				inBuff.close();
				outBuff.close();
				output.close();
				input.close();				
				break;
			}else{
				TimeUnit.MILLISECONDS.sleep(1000);
			} 
		}
	}
	
	public static  String[] readLines(File sourceFile, String encodeType) throws IOException {
		//int bufferSize = 48 * 1024 ;   
		FileInputStream   inStream = new FileInputStream(sourceFile);
		InputStreamReader inRead   = new InputStreamReader(inStream,encodeType);
		
        BufferedReader bufferedReader = new BufferedReader(inRead);
        List<String> lines = new ArrayList<String>();
        String line = null;
        while ((line = bufferedReader.readLine()) != null) {
            lines.add(line);
        }
      
        bufferedReader.close();
        inRead.close();
        inStream.close();
        return lines.toArray(new String[lines.size()]);
    }
	
	public static void rtf2txt(String inputPath,String outPath) throws IOException, BadLocationException, InterruptedException{
		String bodyText = null;
		File sourceFile=new File(inputPath);
        DefaultStyledDocument styledDoc = new DefaultStyledDocument();
        while(true){
        	if(sourceFile.renameTo(sourceFile)){
        		try {
                	InputStream is = new FileInputStream(sourceFile);
                	InputStreamReader inputStreamReader = new InputStreamReader(is,"utf-8");
                    new RTFEditorKit().read(inputStreamReader, styledDoc, 0); //ISO8859_1 
                    byte[] b2=styledDoc.getText(0, styledDoc.getLength()).getBytes("ISO8859_1");
                    bodyText = new String(b2,"GBK");    //提取文本
                    is.close();
                    inputStreamReader.close();
                    FileOutputStream fos = new FileOutputStream(outPath); 
                    OutputStreamWriter osw = new OutputStreamWriter(fos, "GBK"); 
                    osw.write(bodyText); 
                    osw.flush(); 
                    fos.close();
                    osw.close();
                } catch (IOException e) {
                	e.printStackTrace();
                   // System.out.println("IO异常"+e.getMessage());
                } catch (BadLocationException e) {
                	e.printStackTrace();           
                }
        		break;
        	}else{
				TimeUnit.MILLISECONDS.sleep(1000);
			} 
        }
	}
	@SuppressWarnings("finally")
	public static String getTXTEncodeType(File sourceFile) throws IOException{
		String code = "gb2312";
		InputStream inputStream = null;
		try{
			inputStream = new FileInputStream(sourceFile);
			byte[] head = new byte[3];
			inputStream.read(head);
			if (head[0] == -1 && head[1] == -2 )
				code = "UTF-16";
			if (head[0] == -2 && head[1] == -1 )
				code = "Unicode";
			if(head[0]==-17 && head[1]==-69 && head[2] ==-65) 
				code = "UTF-8";   			
		} catch (Exception e) {
        	e.printStackTrace();           
        }finally{
        	inputStream.close();
        	return code;
        }		
	}
	
	  /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
    	int count=0,colNum=0;
    	try{
			Pattern pattern03 = Pattern.compile("\t{1}");
			String[] patternStr = null;	    
			File sourceFile = new File("d:\\7-27.exp");
			if(!sourceFile.exists()){
				System.out.println("okk");	
			}
			String encodeType=FileOptionUtil.getTXTEncodeType(sourceFile);
			String[] lines = FileOptionUtil.readLines(sourceFile,encodeType);
		
			for (String line : lines) {		
				line = line.replaceAll("		", "	");
				patternStr = pattern03.split(line);
				for(colNum=0; colNum<patternStr.length; colNum++){
			        System.out.print(ParseUtil.getObjStr(patternStr[colNum]));
			        System.out.print(" ");			
				}
				System.out.println("");	    
				count=count+1;
			}			
		}catch (Exception e){
			e.printStackTrace();			
			//System.out.println(e.getMessage());
		}	
    }
}
