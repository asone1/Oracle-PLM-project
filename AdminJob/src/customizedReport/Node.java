package customizedReport;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;



public class Node {
	// String label;
	HashMap<Integer, ArrayList<String>> count_Data = new HashMap<Integer, ArrayList<String>>();
	int count = 0;
	

	void setValue(String Status, String Criteria, String CriteriaAPI, String Approvers) {
		
		ArrayList<String> dataList = new ArrayList<String>(4);
		dataList.add(Status);
		dataList.add(Criteria);
		dataList.add(CriteriaAPI);
		dataList.add(Approvers);		
		count_Data.put(count, dataList);
		++count;
		
	}

	public int print(String Department, HSSFSheet sheet, int rowCount) {

		// row0 is row head, created by tableMain.print()

		
			for(int id : count_Data.keySet()) {
				HSSFRow row = sheet.createRow((short) rowCount);
				ArrayList<String> dataList = count_Data.get(id);
				row.createCell(0).setCellValue(Department);
				row.createCell(1).setCellValue(dataList.get(0));
				row.createCell(2).setCellValue(dataList.get(1));
				row.createCell(3).setCellValue(dataList.get(2));
				row.createCell(4).setCellValue(dataList.get(3));
				++rowCount;
			}
			
		return rowCount;

	}
}
