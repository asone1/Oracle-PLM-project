package autoNumber;

import java.time.LocalDate;

import com.agile.api.*;
import com.agile.px.*;

public class PCR_AutoNum implements ICustomAutoNumber {

	private static final String changeType = "PCR";
	// �y�������X�X
	private static int padd4number = 3;

	// ����榡�@�߬� XX�A�Ҧp 09��/10��Aetc
	private static String last_UpdatedMonth = null;
	// �~�榡�@�߬� XX�A�Ҧp (20)20�~�Aetc
	private static String last_UpdatedYear = null;
	private static int next_num = 1;
	int updateLastMonthData = 0;

	public ActionResult getAutoNumber(IAgileSession session, INode actionNode) {

		LocalDate now = LocalDate.now();
		String currentMonth = AutoNumCommonMethod.strAddZero(String.valueOf(now.getMonthValue()), 2);
		String currentYear = String.valueOf(now.getYear()).substring(2, 4);
		String newNumber = "000";

		if (currentMonth == null | currentYear == null) {
			return new ActionResult(ActionResult.EXCEPTION, new Exception("���p��PLM�޲z��"));
		} else {

			// �즸�W�����Ȭ�null�A���X�{�b�ثe��autonumber�A���X�U�@�Ӹ��X
			if (last_UpdatedMonth == null | last_UpdatedYear == null) {
				last_UpdatedMonth = currentMonth;
				last_UpdatedYear = currentYear;
				next_num = AutoNumCommonMethod.setNumber(changeType, padd4number, currentMonth,currentYear , session);
			} else {

				// �p�G���{�����檺����A�P�ثe�t�ΤW���������@��(�i�ର����B������)�A�h�N�ɶ������start number��s��0
				if (!currentMonth.equals(last_UpdatedMonth) | !currentYear.equals(last_UpdatedYear)) {
					last_UpdatedMonth = currentMonth;
					last_UpdatedYear = currentYear;
					next_num = 1;

				}
			}

		}

		// �զX��autoNumeber e.g.PCR-1902001
		newNumber = changeType + "-" + last_UpdatedYear + last_UpdatedMonth
				+ AutoNumCommonMethod.strAddZero(String.valueOf(next_num), padd4number);

		// �����@���W�[1
		++next_num;
		return new ActionResult(ActionResult.STRING, newNumber);
	}

}
