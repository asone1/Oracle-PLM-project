package autoNumber;

import java.time.LocalDate;

import com.agile.api.IAgileSession;
import com.agile.api.INode;
import com.agile.px.ActionResult;
import com.agile.px.ICustomAutoNumber;

public class MJR_AutoNum implements ICustomAutoNumber {

	private static final String changeType = "MJR";
	// 流水號有幾碼
	private static int padd4number = 3;

	// 月份格式一律為 XX，例如 09月/10月，etc
	private static String last_UpdatedMonth = null;
	// 年格式一律為 XX，例如 (20)20年，etc
	private static String last_UpdatedYear = null;
	private static int next_num = 1;

	static String[] getLastMonthData() {
		String array[] = {last_UpdatedMonth,last_UpdatedYear,String.valueOf(next_num)};
		return array;
	}
	
	public ActionResult getAutoNumber(IAgileSession session, INode actionNode) {

		LocalDate now = LocalDate.now();
		String currentMonth = AutoNumCommonMethod.strAddZero(String.valueOf(now.getMonthValue()), 2);
		String currentYear = String.valueOf(now.getYear()).substring(2, 4);
		String newNumber = "000";

		if (currentMonth == null | currentYear == null) {
			return new ActionResult(ActionResult.EXCEPTION, new Exception("請聯絡PLM管理員"));
		} else {

			// 初次上限此值為null，取出現在目前的autonumber，取出下一個號碼
			if (last_UpdatedMonth == null | last_UpdatedYear == null) {
				next_num = AutoNumCommonMethod.setNumber(changeType, padd4number, currentMonth, currentYear, session);
				last_UpdatedMonth = currentMonth;
				last_UpdatedYear = currentYear;
			} else {

				// 如果為程式執行的月份，與目前系統上的紀錄不一樣(可能為次月、次次月)，則將時間跟取號start number更新為0
				if (!currentMonth.equals(last_UpdatedMonth) | !currentYear.equals(last_UpdatedYear)) {
					MJR_LastAutoNum.update(last_UpdatedMonth, last_UpdatedYear, next_num);
					next_num = 1;
					last_UpdatedMonth = currentMonth;
					last_UpdatedYear = currentYear;
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
