package com.yn.spc.util;

import java.beans.PropertyVetoException;
import java.sql.Connection;
import java.sql.SQLException;

import com.mchange.v2.c3p0.ComboPooledDataSource;
import com.mchange.v2.c3p0.DataSources;

public final class ConnectionManager {  
	  
    private static ConnectionManager instance;  
  
    private ComboPooledDataSource ds;  
  
    public ComboPooledDataSource getDs() {
		return ds;
	}

	private ConnectionManager(String url, String userName,
			String password) throws PropertyVetoException {
          
        ds = new ComboPooledDataSource();  

        ds.setDriverClass("oracle.jdbc.driver.OracleDriver");  
        ds.setJdbcUrl(url);  
        ds.setUser(userName);  
        ds.setPassword(password);
        ds.setAutoCommitOnClose(true);
    }  
  
    public static final ConnectionManager getInstance(String url, String userName,
			String password){
        if (instance == null) {  
            try {  
                instance = new ConnectionManager(url,userName,password);  
            } catch (Exception e) {  
                e.printStackTrace();  
            }  
        }  
        return instance;  
    }  
  
    public synchronized final Connection getConnection() {  
        try {  
            return ds.getConnection();  
        } catch (SQLException e) {  
            e.printStackTrace();  
        }  
        return null;  
    }  
    
  
    protected void finalize() throws Throwable {  
        DataSources.destroy(ds); //关闭datasource  
        super.finalize();  
    }

	
  
}  