package event;

import java.io.PrintWriter;
import java.io.StringWriter;
import java.sql.SQLException;
import static event.AEF_COMM_AGILE_FUNCTION.chkEnv;
import com.agile.api.*;
import com.agile.px.*;

public class AddPPM implements IEventAction {

	@Override
	public EventActionResult doAction(IAgileSession arg0, INode arg1, IEventInfo arg2) {


		if (!chkEnv()) {
			return new EventActionResult(arg2, new ActionResult(ActionResult.NORESULT, null));
		}

		IUpdateEventInfo info = null;
		Integer subclassId = null;
		String newNumber = null;
		try {
			//���\save as��create�ʧ@�I�s�{��
			switch (arg2.getEventType()) {
			case EventConstants.EVENT_CREATE_OBJECT:
				info = (ICreateEventInfo) arg2;
				subclassId = ((ICreateEventInfo) info).getNewSubclassId();
				newNumber = ((ICreateEventInfo) info).getNewNumber();
				break;
			case EventConstants.EVENT_SAVE_AS_OBJECT:
				info = (ISaveAsEventInfo) arg2;
				subclassId = ((ISaveAsEventInfo) info).getNewSubclassId();
				newNumber = ((ISaveAsEventInfo) info).getNewNumber();
				break;
			default:
				return new EventActionResult(arg2,
						new ActionResult(ActionResult.EXCEPTION, new Exception("�q��admin�A�u��z�Lsave as��create�ʧ@�I�s")));
			
			}

			//listID 224048��DCN�ҥ� Page Three��� ���ݱM�צW��
			int listID = 224048;
			String newProgramName = "";// ex:AKT_RTF500_GNSS F9P

			// ���o�s�إߪ�����
			if (subclassId == null | newNumber == null) {
				return new EventActionResult(arg2,
						new ActionResult(ActionResult.EXCEPTION, new Exception("ID��null�A���p��PLM ADMIN")));
			} else {
				IProgram newProgram = (IProgram) arg0.getObject(subclassId, newNumber);
				newProgramName = newProgram.getValue(ProgramConstants.ATT_GENERAL_INFO_NAME).toString();

				// �۰ʥ[���
				// ���o��LIST��쪫��
				IAdmin admin = arg0.getAdminInstance();
				IListLibrary listLib = (IListLibrary) admin.getListLibrary();
				IAdminList list = listLib.getAdminList(listID);
				IAgileList listVals = list.getValues();

				// �Nprogram�W�٥[�J��list
				listVals.addChild(newProgramName);
				list.setValues(listVals);

				return new EventActionResult(arg2,
						new ActionResult(ActionResult.STRING, "�۰ʩ�LIST�i" + listID + "�j�[�J" + newProgramName));
			}

		} catch (Exception e) {
			e.printStackTrace();
			return new EventActionResult(arg2, new ActionResult(ActionResult.EXCEPTION, e));
		}

	}
}
