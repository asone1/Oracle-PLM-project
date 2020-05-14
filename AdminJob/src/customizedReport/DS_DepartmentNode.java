package customizedReport;

import java.util.ArrayList;
import java.util.HashMap;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class DS_DepartmentNode {
	HashMap<String, Node> nameList = new HashMap<String, Node>();

	static int rowCount=2;
	
	public void build(String Depart, String Status, String Criteria, String CriteriaAPI, String Approvers) {

		if (!nameList.containsKey(Depart)) {
			Node newNode =new Node();
			newNode.setValue(Status, Criteria, CriteriaAPI, Approvers);
			nameList.put(Depart, newNode);
			

		} else {

			Node temp = nameList.get(Depart);
			temp.setValue(Status, Criteria, CriteriaAPI, Approvers);

		}

	}

	public void printReport(HSSFWorkbook workbook, String workflowName) {
		
		HSSFSheet sheet = workbook.createSheet(workflowName);
		sheet.setColumnWidth(0, 2000);
		sheet.setColumnWidth(1, 5000);
		sheet.setColumnWidth(2, 5000);
		sheet.setColumnWidth(3, 5000);
		sheet.setColumnWidth(4, 7000);
		HSSFRow rowTopic = sheet.createRow((short) 0);
		rowTopic.createCell(0).setCellValue(workflowName+"簽核清單");
		HSSFRow rowhead = sheet.createRow((short) 1);
		rowhead.createCell(0).setCellValue("事業處");
		rowhead.createCell(1).setCellValue("Status");
		rowhead.createCell(2).setCellValue("Criteria");
		rowhead.createCell(3).setCellValue("CriteriaAPI");
		rowhead.createCell(4).setCellValue("Approvers");
		
		
		for (String Department : nameList.keySet()) {

			int rowCount = DS_DepartmentNode.rowCount;
			if (Department == null || sheet == null || nameList.get(Department) == null) {
			} else {
				DS_DepartmentNode.rowCount = nameList.get(Department).print(Department, sheet,rowCount);
			}
		}
		
		nameList = new HashMap<String, Node>();
		DS_DepartmentNode.rowCount=2;
	}

}
