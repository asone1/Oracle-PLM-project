package autoNumber;

import java.time.LocalDate;

import com.agile.api.*;
import com.agile.px.*;

public class AutoNumCommonMethod {
	public static String strAddZero(String str, int len) {
		int strLen = str.length();
		if (strLen < len) {
			while (strLen < len) {
				StringBuffer sb = new StringBuffer();
				sb.append("0").append(str);// ����0
				str = sb.toString();
				strLen = str.length();
			}
		}
		return str;
	}
	
	/**********************
	 * �ΨӨ��X�ثe����A�����U�@��number�A�즸�W�u�~�|�Ψ�
	 **********************/
	public static int setNumber(String changeType, int padd4number,String currentMonth,
			String currentYear, IAgileSession session) {
		try {
			IAgileClass partSubclass = session.getAdminInstance().getAgileClass(changeType);
			IAutoNumber[] autoNumbers = partSubclass.getAutoNumberSources();
			
			int nextNumber=0;
			String autoNumSourceName = changeType + "-" + currentYear + currentMonth;
			
			//���{�b�Ϊ�autoNumber source
			for (IAutoNumber a : autoNumbers) {
				if (a.getName().toString().trim().equalsIgnoreCase(autoNumSourceName)) {
					String current_auto_number = a.getNextNumber();
					nextNumber = Integer.parseInt(current_auto_number
							.substring(current_auto_number.length() - padd4number, current_auto_number.length()));
				}
			}
			return nextNumber;
		} catch (APIException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		return -1;
	}

	
}
