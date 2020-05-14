package autoNumber;

import java.time.LocalDate;

import com.agile.api.IAgileSession;
import com.agile.api.INode;
import com.agile.px.ActionResult;
import com.agile.px.ICustomAutoNumber;

public class MJR_LastAutoNum implements ICustomAutoNumber {

	private static final String changeType = "MJR";
	// �y�������X�X
	private static int padd4number = 3;

	// ����榡�@�߬� XX�A�Ҧp 09��/10��Aetc
	private static String last_UpdatedMonth = null;
	// �~�榡�@�߬� XX�A�Ҧp (20)20�~�Aetc
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
			return new ActionResult(ActionResult.EXCEPTION, new Exception("���p��PLM�޲z��"));
		} else {

			// �즸�W�����Ȭ�null�A���X�{�b�ثe��autonumber�A���X�U�@�Ӹ��X
			if (last_UpdatedMonth == null | last_UpdatedYear == null) {
				next_num = AutoNumCommonMethod.setNumber(changeType, padd4number, lastMonth, currentYear, session);
				last_UpdatedMonth = lastMonth;
				last_UpdatedYear = currentYear;
			} else {

				// �p�G�{�����檺����A�P�ثe�t�ΤW���������@��(�i�ର����B������)�A�h��s���
				if (!lastMonth.equals(last_UpdatedMonth) | !currentYear.equals(last_UpdatedYear)) {
					
					//���]MJR���Q����L�Adirty data�������
					if(dirtyMonth == lastMonth) {
						last_UpdatedMonth=dirtyMonth;
						last_UpdatedYear=dirtyYear;
						next_num=dirtyNextNum;
						dirtyNextNum=-1;
					}
					
					//�YMJR�S�Q����A��X�ثe���
					else if (dirtyNextNum==-1) {
						//get data from current MJR
						String temp [] = MJR_AutoNum.getLastMonthData();
						//��MJR������ �{������������e�@�Ӥ� �~��sLAST���
						if(temp[0] ==lastMonth ) {
							last_UpdatedMonth = temp[0];
							last_UpdatedYear= temp[1];
							next_num= Integer.parseInt(temp[2]);
						}
						
					}
					
					//MJR������Ƭ��W�W�몺�A�N��W��S�����C�]����s��AUTO���W��Anumber�q0�}�l
					else {
						last_UpdatedMonth = lastMonth;
						last_UpdatedYear= currentYear;
						next_num= 0;
						
					}
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
