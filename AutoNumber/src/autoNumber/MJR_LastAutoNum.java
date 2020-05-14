package autoNumber;

import java.time.LocalDate;

import com.agile.api.IAgileSession;
import com.agile.api.INode;
import com.agile.px.ActionResult;
import com.agile.px.ICustomAutoNumber;

public class MJR_LastAutoNum implements ICustomAutoNumber {

	private static final String changeType = "MJR";
	// 流水號有幾碼
	private static int padd4number = 3;

	// 月份格式一律為 XX，例如 09月/10月，etc
	private static String last_UpdatedMonth = null;
	// 年格式一律為 XX，例如 (20)20年，etc
	private static String last_UpdatedYear = null;
	private static int next_num = 1;

	private static String dirtyMonth= null;
	private static String dirtyYear= null;
	private static int dirtyNextNum= -1;
	
	static void update(String thisMonth, String thisYear, int nextNumber) {
		dirtyMonth = thisMonth;
		dirtyYear = thisYear;
		dirtyNextNum = nextNumber;
		
	}
	
	public ActionResult getAutoNumber(IAgileSession session, INode actionNode) {

		LocalDate now = LocalDate.now();
		String lastMonth = AutoNumCommonMethod.strAddZero(String.valueOf(now.getMonthValue()-1), 2);
		String currentYear="";
		String newNumber = "000";

		if(lastMonth=="00") {
			lastMonth ="12";
			currentYear = String.valueOf(Integer.parseInt(lastMonth) -1);
		}else {
			currentYear = String.valueOf(now.getYear()).substring(2, 4);
		}
		
		if (lastMonth == null | currentYear == null) {
			return new ActionResult(ActionResult.EXCEPTION, new Exception("請聯絡PLM管理員"));
		} else {

			// 初次上限此值為null，取出現在目前的autonumber，取出下一個號碼
			if (last_UpdatedMonth == null | last_UpdatedYear == null) {
				next_num = AutoNumCommonMethod.setNumber(changeType, padd4number, lastMonth, currentYear, session);
				last_UpdatedMonth = lastMonth;
				last_UpdatedYear = currentYear;
			} else {

				// 如果程式執行的月份，與目前系統上的紀錄不一樣(可能為次月、次次月)，則更新資料
				if (!lastMonth.equals(last_UpdatedMonth) | !currentYear.equals(last_UpdatedYear)) {
					
					//假設MJR先被執行過，dirty data應有資料
					if(dirtyMonth == lastMonth) {
						last_UpdatedMonth=dirtyMonth;
						last_UpdatedYear=dirtyYear;
						next_num=dirtyNextNum;
						dirtyNextNum=-1;
					}
					
					//若MJR沒被執行，找出目前資料
					else if (dirtyNextNum==-1) {
						//get data from current MJR
						String temp [] = MJR_AutoNum.getLastMonthData();
						//當MJR紀錄為 程式執行紀錄的前一個月 才更新LAST資料
						if(temp[0] ==lastMonth ) {
							last_UpdatedMonth = temp[0];
							last_UpdatedYear= temp[1];
							next_num= Integer.parseInt(temp[2]);
						}
						
					}
					
					//MJR有的資料為上上月的，代表上月沒有表單。因此更新此AUTO為上月，number從0開始
					else {
						last_UpdatedMonth = lastMonth;
						last_UpdatedYear= currentYear;
						next_num= 0;
						
					}
				}
			}

		}
		// 組合成autoNumeber e.g.PCR-1902001
		newNumber = changeType + "-" + last_UpdatedYear + last_UpdatedMonth
				+ AutoNumCommonMethod.strAddZero(String.valueOf(next_num), padd4number);

		// 取號一次增加1
		++next_num;
		return new ActionResult(ActionResult.STRING, newNumber);
	}
}
