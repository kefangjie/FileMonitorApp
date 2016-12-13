/*
 * CVS Info
 * 
 * $Id: HibernateUtil.java,v 1.4 2012/11/23 06:27:08 hmei Exp $
 * 
 * $Log: HibernateUtil.java,v $
 * Revision 1.4  2012/11/23 06:27:08  hmei
 * bugfix
 *
 * Revision 1.3  2012/11/21 08:23:47  hmei
 * Add closeAll
 *
 * Revision 1.2  2012/11/19 04:32:01  hmei
 * Add closeAll
 *
 * Revision 1.1  2012/05/30 02:24:37  kclai
 * GEM - JunLin
 *
 * Revision 1.2  2007/01/24 04:18:53  khtse
 * avoid session is closed
 *
 * Revision 1.1  2006/09/28 02:42:11  ktfu
 * added from weberpjbossdev
 *
 * Revision 1.2  2006/08/30 08:26:41  slam
 * rename hibernate/SessionFactory to SessionFactory
 *
 * Revision 1.1  2006/06/29 09:35:12  slam
 * initial version
 *
 * Revision 1.1.1.1  2006/06/08 03:36:07  slam
 * ecard project on hibernate
 *
 * Revision 1.1  2006/06/05 08:00:39  slam
 * initial version
 *
 * 
 */
package com.yn.spc.util;



import java.io.BufferedInputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.SQLException;
import java.util.Properties;

import org.hibernate.FlushMode;
import org.hibernate.HibernateException;
import org.hibernate.MappingException;
import org.hibernate.Session;
import org.hibernate.SessionFactory;
import org.hibernate.Transaction;
import org.hibernate.cfg.Configuration;
import org.hibernate.service.ServiceRegistry;
import org.hibernate.service.ServiceRegistryBuilder;


import com.mchange.v2.c3p0.ComboPooledDataSource;
import com.yn.spc.exceptions.ServiceLocatorException;
import com.yn.spc.ui.SPCUI;

public class DBOptionUtil {
	
	public static SessionFactory sessionFactory; 
	
	public static ConnectionManager cm;  
	
	public static final ThreadLocal<Connection> connectionHolder=new ThreadLocal<Connection>();
	
	public static final ThreadLocal<Session> session = new ThreadLocal<Session>(); 
	
	public static final ThreadLocal<Transaction> transaction = new ThreadLocal<Transaction>();
	
	public static final String HIBERNATE_PATH=System.getProperty("user.dir") + "/conf/hibernate.properties";

	public static void dataSourceConf(String url,String userName,String password) throws MappingException, IOException{
		InputStream in =null;
		Properties p= new Properties();
		try{
			in = new BufferedInputStream(new FileInputStream(HIBERNATE_PATH));
			p.load(in);
		}catch(Exception e){
			e.printStackTrace();
			SPCUI.log.append(e.getMessage());
		}finally{
			in.close();
		}
		Configuration configuration=new Configuration().addProperties(p);	

		
		configuration.setProperty( "hibernate.connection.url" , url);
		configuration.setProperty( "hibernate.connection.username" ,userName );
		configuration.setProperty( "hibernate.connection.password" , password);
	
		ServiceRegistry  sr = new ServiceRegistryBuilder().applySettings(configuration.getProperties()).buildServiceRegistry(); 
		
		sessionFactory = configuration.buildSessionFactory(sr);
		
		cm= ConnectionManager.getInstance(url,userName,password);
		
		if(cm.getDs() instanceof ComboPooledDataSource){
			SPCUI.log.append("数据库连接成功！\n");
			p.setProperty("hibernate.connection.url", url); 
			p.setProperty("hibernate.connection.username",userName);
			p.setProperty("hibernate.connection.password", password);
			FileOutputStream fos = new FileOutputStream(HIBERNATE_PATH); 
			try{
				p.store(fos, "Copyright (c) TT "); 
			}catch(Exception e){
				e.printStackTrace();
				SPCUI.log.append(e.getMessage());
			}finally{
				fos.close();
			}
		}
	}
	
	//保存数据库信息
	public static void saveDataSourceInfo(Properties p){
		
	}
	
	public static Session getSession() throws HibernateException {
		Session s = (Session)session.get();
		if (s == null || !s.isOpen()) {
			s = null;
			try{
				s = sessionFactory.openSession();
				s.setFlushMode(FlushMode.AUTO);
				session.set(s);
			}catch(Exception e){
				SPCUI.log.append("数据库未连接,请先连接数据库\n");
			}			
		}
		return s;
	}
	
	public static void closeSession() throws HibernateException {
		Session s = (Session)session.get();
		session.set(null);
		if (s!=null && s.isOpen()) {
			s.close();
		}
	}
	
	public static void beginTransaction() {
		Transaction tx = (Transaction)transaction.get();
		if (tx == null) {
			tx = getSession().beginTransaction();
			transaction.set(tx);
		}
	}
	
	public static void commitTransaction() {
		Transaction tx = (Transaction)transaction.get();
		try {
			if (tx != null && !tx.wasCommitted() && !tx.wasRolledBack()) {
				getSession().flush();
				tx.commit();
			}
			transaction.set(null);
		} catch(HibernateException ex) {
			rollbackTransaction();
			throw ex;
		}
	}
	
	public static void rollbackTransaction() {
		Transaction tx = (Transaction)transaction.get();
		try {
			if (tx != null && !tx.wasCommitted() && !tx.wasRolledBack()) {
				tx.rollback();
			}
		} catch (HibernateException ex){
			throw ex;
		} finally {
			closeSession();
		}
	}
	
	public static void closeThreadLocalConnection(){
        Connection conn = connectionHolder.get();
        if (conn != null){
            try{
                conn.close();
                connectionHolder.remove();
            }catch(SQLException e){
                    e.printStackTrace();
            }
        }
	}
	
	public static Connection getConnection() throws HibernateException {
		Connection conn=null;
		try {
			 conn = cm.getConnection();
		} catch (Exception e) {
			e.printStackTrace();
		}
		return conn;
	}
	
	public static Connection getThreadLocalConnection() throws HibernateException, ServiceLocatorException, SQLException {
	       Connection conn = connectionHolder.get();
           if (conn == null){
               conn = cm.getConnection();
               connectionHolder.set(conn);
           }       
           return conn;
	}
	
	public static void close() throws Throwable{
		if(cm!=null)
			cm.finalize();
		connectionHolder.remove();
		session.remove();
		transaction.remove();
	}
	
}
