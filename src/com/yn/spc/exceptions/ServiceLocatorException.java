/* 
 * CVS Info: 
 * 
 * $Id: ServiceLocatorException.java,v 1.1 2012/05/30 02:24:54 kclai Exp $ 
 * 
 * $Log: ServiceLocatorException.java,v $
 * Revision 1.1  2012/05/30 02:24:54  kclai
 * GEM - JunLin
 *
 * Revision 1.3  2007/04/25 05:36:58  slam
 * added serialVersionUID
 *
 * Revision 1.2  2006/05/25 06:16:22  ktfu
 * *** empty log message ***
 *
 * Revision 1.1  2006/04/11 01:58:53  ktfu
 * *** empty log message ***
 *
 * Revision 1.1  2006/04/11 01:57:53  ktfu
 * *** empty log message ***
 *
 * Revision 1.1  2006/04/10 09:10:02  ktfu
 * *** empty log message ***
 *
 * Revision 1.2  2006/04/06 06:50:13  ktfu
 * Initial upload
 *
 *
 */ 
package com.yn.spc.exceptions;

import java.io.Serializable;

public class ServiceLocatorException extends RuntimeException implements Serializable{
	private static final long serialVersionUID = -7513807849076126614L;

	public ServiceLocatorException() {
        super();
    }

    public ServiceLocatorException(String message) {
        super(message);
    }

    public ServiceLocatorException(Throwable cause) {
        super(cause);
    }

    public ServiceLocatorException(String message, Throwable cause) {
        super(message, cause);
    }
}
