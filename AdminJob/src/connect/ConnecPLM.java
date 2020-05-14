package connect;

import java.net.InetAddress;
import java.net.UnknownHostException;
import java.util.HashMap;

import com.agile.api.APIException;
import com.agile.api.AgileSessionFactory;
import com.agile.api.IAgileSession;
import com.agile.api.ICriteria;

public class ConnecPLM {
	public static IAgileSession logInAsAdmin() {
		AgileSessionFactory instance = null;
		IAgileSession myServer = null;
		HashMap params = new HashMap();
		 final String DEV_ADDR = "http://10.0.88.74:7001/Agile";
		 final String PROD_ADDR = "http://10.0.88.61:7001/Agile";
		try {
			
			
			String loginADDR = DEV_ADDR;
			//instance = AgileSessionFactory.getInstance(loginADDR);
			
			// prod
			instance = AgileSessionFactory.getInstance(loginADDR);
			if(loginADDR.equals(PROD_ADDR)) {
				System.out.println("正式區");
			}else {
				System.out.println("TEST");
			}
			params.put(AgileSessionFactory.USERNAME, "plmadmin");
			params.put(AgileSessionFactory.PASSWORD, "cj06xj/6");
			myServer = instance.createSession(params);
			
			
		

		} catch (APIException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		return myServer;

	}
	
	public static String getCriteria_noNULL(ICriteria c) {
		String s = null;
		try {
			s = c.getCriteria();
		} catch (NullPointerException e) {
			s = "找不到criteria";
			return s;
		} catch (APIException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return s;
	}

}
