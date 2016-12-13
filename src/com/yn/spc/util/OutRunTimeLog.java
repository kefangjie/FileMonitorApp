package com.yn.spc.util;

import com.yn.spc.ui.SPCUI;

public  class OutRunTimeLog {
	//日志输出同步方法
	public static synchronized  void showLog(String msg){

		if(SPCUI.log.getLineCount()>=2000) {
			SPCUI.log.setText("");
		}
		SPCUI.log.append(msg);		
	}
}
