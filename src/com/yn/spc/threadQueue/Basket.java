package com.yn.spc.threadQueue;

import java.util.concurrent.BlockingQueue;
import java.util.concurrent.LinkedBlockingQueue;

/**
 * 文件监测发现与读取
 *  
 * @author jxke
 * @version 1.0 
 */

    /**
     * 
     * 定义装苹果的篮子
     * 
     */
  public class Basket {
     
        public static  BlockingQueue<String> basketQueue = new LinkedBlockingQueue<String>(50000);
        
		private String sourceFilePath;

		private String pathValue;
		
		public Basket(String sourceFilePath, String pathValue) {
			super();
			this.sourceFilePath = sourceFilePath;
			this.pathValue = pathValue;
		}
       
        public void produce() throws InterruptedException {         
        	basketQueue.put(this.sourceFilePath+"=#="+this.pathValue);
        }

   
        public String consume() throws InterruptedException {           
            return basketQueue.take();
        }
        
		public String getSourceFilePath() {
			return sourceFilePath;
		}

		public void setSourceFilePath(String sourceFilePath) {
			this.sourceFilePath = sourceFilePath;
		}

		public String getPathValue() {
			return pathValue;
		}

		public void setPathValue(String pathValue) {
			this.pathValue = pathValue;
		}
   }





