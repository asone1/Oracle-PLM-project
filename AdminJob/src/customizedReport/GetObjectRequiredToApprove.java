package customizedReport;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.agile.api.APIException;
import com.agile.api.IAgileSession;
import com.agile.api.ICriteria;
import com.agile.api.INode;
import com.agile.api.IProperty;
import com.agile.api.IRow;
import com.agile.api.ITable;
import com.agile.api.IUserGroup;
import com.agile.api.NodeConstants;
import com.agile.api.UserGroupConstants;
import connect.*;

public class GetObjectRequiredToApprove {
	public static final String no_approver = "no approvers";
	public static final String title = "須簽核清單";

	public static INode toNode(Object o) {
		return (INode) o;
	}

	public static boolean ifEmpty(String s) {
		if (s.trim().isEmpty() || s == null) {
			return true;
		} else {
			return false;
		}
	}

	public static HSSFWorkbook getReport(IAgileSession myServer, String ReplacedApprover) {

		HSSFWorkbook workbook = new HSSFWorkbook();
		HSSFSheet sheet = workbook.createSheet(title);
		sheet.setColumnWidth(0, 2000);
		sheet.setColumnWidth(1, 5000);
		sheet.setColumnWidth(2, 5000);
		sheet.setColumnWidth(3, 5000);
		sheet.setColumnWidth(4, 7000);
		HSSFRow rowTopic = sheet.createRow((short) 0); // rowCount ++
		rowTopic.createCell(0).setCellValue(title);
		HSSFRow rowhead = sheet.createRow((short) 1); // rowCount ++
		rowhead.createCell(0).setCellValue("表單種類");
		rowhead.createCell(1).setCellValue("workflow");
		rowhead.createCell(2).setCellValue("表單站點");
		rowhead.createCell(3).setCellValue("Criteria");
		rowhead.createCell(4).setCellValue("CriteriaAPI");
		rowhead.createCell(5).setCellValue("目前簽核者");
		int rowCount = 2;

		DS_DepartmentNode myReport = new DS_DepartmentNode();
		try {
			Object[] workflows = myServer.getAdminInstance().getNode(NodeConstants.NODE_AGILE_WORKFLOWS).getChildren();
			if (workflows == null) {
				System.out.println("null");
			}
			// 所有workflow
			for (Object wf : workflows) {
				INode wf_node = (INode) wf;
				String workflowStr = wf_node.getName().toString();
			
				// 找所有正在使用的workflow
				if (!workflowStr.toLowerCase().startsWith("x")) {
					Object[] children = wf_node.getChildren();
					System.out.println(wf_node.getName());
					// wf_node has many children, , which contains status
					// get node:[Status List]
					INode wf_general = (INode) children[0];
					INode statuses = (INode) children[1];
					Object[] generalObj = wf_general.getChildren();
					Object[] status = statuses.getChildren();

					int countWF =0;
					INode gerralNode = (INode) generalObj[countWF];
					IProperty[] wfProperties = gerralNode.getProperties();
					String ObjectType = wfProperties[2].getValue().toString();
					
	
					// for every status e.g.intitiate, internal review
					for (int i = 0; i < status.length; ++i) {
						//general頁籤一起拿
												
						//status頁籤
						INode wfStatus = (INode) status[i];
						String status_str = "";
						String criteria_str = "";
						String criteriaAPI_str = "";
						String approvers_str = no_approver;
						List<String> approver_List = new ArrayList<String>();// 只供檢查approvers裡面是否包含user group

						// status_node has two children:[0]Status-Specific Properties
						// node；[1]:Criteria-Specific Properties node
						Object[] status_node = wfStatus.getChildren();
						Object[] N_Criteria_Specific = toNode(status_node[1]).getChildren();

						// in each end criteria-specific node, get its properties: criteriaAPI and
						// approvers
						for (Object o : N_Criteria_Specific) {
							IProperty[] pp = (IProperty[]) toNode(o).getProperties();

							

							// excel cell[3]:approvers
							if (!ifEmpty(pp[10].getValue().toString())) {
								
								for (IProperty ppp : pp) {
									if (ppp.getName().equalsIgnoreCase("API Name")) {
										// ++count;
										

										// excel cell:criteria workflow
										criteriaAPI_str = ppp.getValue().toString();

										INode criteriaLib = myServer.getAdminInstance()
												.getNode(NodeConstants.NODE_CRITERIA_LIBRARY);
										ICriteria criteria = (ICriteria) criteriaLib.getChild(ppp.getValue().toString());

										// excel cell:criteria attributes
										criteria_str = ConnecPLM.getCriteria_noNULL(criteria);
										criteria_str = criteria_str.replaceAll("equal to", "為").replaceAll("contains", "包含").replaceAll("and", "且").replaceAll("or", "或");
									}
								}
								
								// excel cell:status
								status_str = wfStatus.getName();

								approvers_str = pp[10].getValue().toString();
								String[] approver_Array = approvers_str.split(";");
								approver_List = Arrays.asList(approver_Array);
								for (String each_approver : approver_List) {
									char approver_1stChar = approvers_str.charAt(0);
									if ((approver_1stChar >= 'a' && approver_1stChar <= 'z')
											|| (approver_1stChar >= 'A' && approver_1stChar <= 'Z')) {

										if (myServer.getObject(IUserGroup.OBJECT_TYPE, approvers_str) != null) {
											IUserGroup userGroup = (IUserGroup) myServer
													.getObject(IUserGroup.OBJECT_TYPE, approvers_str);
											Iterator it = null;
											ITable ug_tab2 = userGroup.getTable(UserGroupConstants.TABLE_USERS);
											it = ug_tab2.iterator();

											// 將user_group的approver註解在後面
											approvers_str += "【" + each_approver + ":";
											while (it.hasNext()) {
												IRow Row = (IRow) (it.next());
												approvers_str += Row.getValue(UserGroupConstants.ATT_USERS_USER_NAME)
														.toString();
											}
											approvers_str += "】";

										}
									} // 找英文開頭的approver

									
									if (approvers_str.contains(ReplacedApprover)) {
										HSSFRow row = sheet.createRow((short) rowCount);
										row.createCell(0).setCellValue(ObjectType);
										row.createCell(1).setCellValue(workflowStr);
										row.createCell(2).setCellValue(status_str);
										row.createCell(3).setCellValue(criteria_str);
										row.createCell(4).setCellValue(criteriaAPI_str);
										row.createCell(5).setCellValue(approvers_str);
										++rowCount;
									}

								} // 所有approver

							} // 找approver的點
						} // 該站點criteria

						++countWF;
					} // 所有站點
				} // 找所有正在使用的workflow
			} // 所有workflow
		} catch (APIException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		return workbook;
	}

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		IAgileSession myServer = ConnecPLM.logInAsAdmin();
		if (myServer != null) {
			System.out.println("connected");
			// Scanner S=new Scanner(System.in)
			String approver = "u5600";
			String filename = "C:/Users/u10087/Desktop/" + approver + "_WF_AppovalReport.xls";
			FileOutputStream fileOut = null;

			// get a file output stream
			try {
				fileOut = new FileOutputStream(filename);
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			HSSFWorkbook workbook = getReport(myServer, approver);

			// write into excel
			try {
				workbook.write(fileOut);
				fileOut.close();
				System.out.println("Your excel file has been generated!");
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			System.out.println("END");
		} else {
			System.out.println("Connection is null");
		}

	}

}
