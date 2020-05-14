package customizedReport;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import com.agile.api.*;

import connect.ConnecPLM;

public class Workflow_criteria_approver {
	public static final String no_approver = "no approvers";

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

	public static String getDpart(String strToDepart) {
		String department = "";
		if (strToDepart.contains("TW2") || strToDepart.contains("JL21") || strToDepart.contains("JL22")
				|| strToDepart.contains("VL21") || strToDepart.contains("VL22") || strToDepart.contains("電磁")) {
			department += "電磁";
		} else if (strToDepart.contains("TW3") || strToDepart.contains("JL26") || strToDepart.contains("JL31")
				|| strToDepart.contains("電源")) {
			department += "電源";
		} else if (strToDepart.contains("TW4") || strToDepart.contains("JL27") || strToDepart.contains("資通")) {
			department += "資通";
		} else if (strToDepart.contains("TW5") || strToDepart.contains("JL5") || strToDepart.contains("光通訊")) {
			department += "光通訊";
		} else if (strToDepart.contains("JL7") || strToDepart.contains("光電")) {
			department += "光電";
		} else if (strToDepart.contains("TW14") || strToDepart.contains("JL14") || strToDepart.contains("採購")) {
			department += "採購";
		} else if (strToDepart.contains("TW6") || strToDepart.contains("JL6") || strToDepart.contains("品保")) {
			department += "品保";
		} else if (strToDepart.contains("CE")) {
			department += "零件工程";
		} else if (strToDepart.contains("TW10") || strToDepart.contains("JL10")) {
			department += "總經理室";
		} else {
			department = no_approver;
		}
		return department;
	}

	

	public static HSSFWorkbook getReport(IAgileSession myServer, String[] workflowList) {

		HSSFWorkbook workbook = new HSSFWorkbook();
		DS_DepartmentNode myReport = new DS_DepartmentNode();

		for (String workflowName : workflowList) {

			try {

				// myServer.disableAllWarnings();

				Object[] workflows = myServer.getAdminInstance().getNode(NodeConstants.NODE_AGILE_WORKFLOWS)
						.getChildren();
				if (workflows == null) {
					System.out.println("null");
				}
				
				for (Object wf : workflows) {
					INode wf_node = (INode) wf;

					// find the specific workflow
					if (wf_node.getName().equalsIgnoreCase(workflowName)) {
						Object[] children = wf_node.getChildren();

						// wf_node has many children, , which contains status
						// get node:[Status List]
						INode statuses = (INode) children[1];
						Object[] status = statuses.getChildren();

						// for every status e.g.intitiate, internal review
						for (int i = 0; i < status.length; ++i) {
							INode wfStatus = (INode) status[i];
							String department_str = "";
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

								for (IProperty ppp : pp) {
									if (ppp.getName().equalsIgnoreCase("API Name")) {
										// ++count;
										// excel cell:status
										status_str = wfStatus.getName();

										// excel cell:criteria workflow
										criteriaAPI_str = ppp.getValue().toString();

										INode criteriaLib = myServer.getAdminInstance()
												.getNode(NodeConstants.NODE_CRITERIA_LIBRARY);
										ICriteria criteria = (ICriteria) criteriaLib
												.getChild(ppp.getValue().toString());

										// excel cell:criteria attributes
										criteria_str = ConnecPLM.getCriteria_noNULL(criteria);

									}
								}

								// excel cell[3]:approvers
								if (!ifEmpty(pp[10].getValue().toString())) {

									approvers_str = pp[10].getValue().toString();
									String[] approver_Array = approvers_str.split(";");
									approver_List = Arrays.asList(approver_Array);
									for (String each_approver : approver_List) {
										char approver_1stChar = approvers_str.charAt(0);
										if ((approver_1stChar >= 'a' && approver_1stChar <= 'z')
												|| (approver_1stChar >= 'A' && approver_1stChar <= 'Z')) {
											System.out.print(each_approver + "\n");

											try {
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
														approvers_str += Row
																.getValue(UserGroupConstants.ATT_USERS_USER_NAME)
																.toString();
													}
													approvers_str += "】";
												}

											} catch (APIException e) {
											}
										}
									}

								}

								department_str = getDpart(criteriaAPI_str);
								if (department_str == no_approver) {
									department_str = getDpart(approvers_str);
								}

								myReport.build(department_str, status_str, criteria_str, criteriaAPI_str,
										approvers_str);
							}
						}
					}

				}

			} catch (APIException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}

			myReport.printReport(workbook, workflowName);
		}
		myServer.close();
		// workbook.close();
		return workbook;
	}

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		IAgileSession myServer = ConnecPLM.logInAsAdmin();
		if (myServer != null) {
			System.out.println("connected");

			String filename = "C:/Users/u10087/Desktop/workflow_approver_report.xls";
			FileOutputStream fileOut = null;
			String[] workflowList = { "DMN", "MJR", "Audit new" };
//		String[] workflowList = { "PNRF SFG_FG", "PNRF Component", "ECO", "Material Approval", "Obsolete-work", "DCN",
//				"MCO", "EOL", "IE Document Release" };

			// get a file output stream
			try {
				fileOut = new FileOutputStream(filename);
			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			// generate report
			HSSFWorkbook workbook = getReport(myServer, workflowList);

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
		}else {
			System.out.println("Connection is null");
		}
	}

}
