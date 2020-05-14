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
			//允許save as或create動作呼叫程式
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
						new ActionResult(ActionResult.EXCEPTION, new Exception("通知admin，只能透過save as或create動作呼叫")));
			
			}

			//listID 224048為DCN所用 Page Three欄位 所屬專案名稱
			int listID = 224048;
			String newProgramName = "";// ex:AKT_RTF500_GNSS F9P

			// 取得新建立的物件
			if (subclassId == null | newNumber == null) {
				return new EventActionResult(arg2,
						new ActionResult(ActionResult.EXCEPTION, new Exception("ID為null，請聯絡PLM ADMIN")));
			} else {
				IProgram newProgram = (IProgram) arg0.getObject(subclassId, newNumber);
				newProgramName = newProgram.getValue(ProgramConstants.ATT_GENERAL_INFO_NAME).toString();

				// 自動加欄位
				// 取得該LIST欄位物件
				IAdmin admin = arg0.getAdminInstance();
				IListLibrary listLib = (IListLibrary) admin.getListLibrary();
				IAdminList list = listLib.getAdminList(listID);
				IAgileList listVals = list.getValues();

				// 將program名稱加入該list
				listVals.addChild(newProgramName);
				list.setValues(listVals);

				return new EventActionResult(arg2,
						new ActionResult(ActionResult.STRING, "自動於LIST【" + listID + "】加入" + newProgramName));
			}

		} catch (Exception e) {
			e.printStackTrace();
			return new EventActionResult(arg2, new ActionResult(ActionResult.EXCEPTION, e));
		}

	}
}
