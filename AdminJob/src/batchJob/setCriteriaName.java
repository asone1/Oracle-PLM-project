package batchJob;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.agile.api.APIException;
import com.agile.api.IAgileSession;
import com.agile.api.ICriteria;
import com.agile.api.INode;
import com.agile.api.NodeConstants;

import connect.ConnecPLM;

public class setCriteriaName {

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		Map<String, String> accordMap = new HashMap<>();
		IAgileSession myServer = ConnecPLM.logInAsAdmin();

		String Infilename = "C:/Users/u10087/Desktop/replaceList.xls";
		FileInputStream fis = null;
		HSSFWorkbook accordBook = null;
		try {
			fis = new FileInputStream(new File(Infilename));
			accordBook = new HSSFWorkbook(fis);
		} catch (FileNotFoundException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		} catch (IOException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		}

		HSSFSheet sheettemp2 = accordBook.getSheetAt(0);
		for (int rowIndex = 0; rowIndex <= sheettemp2.getLastRowNum(); rowIndex++) {
			Row row = sheettemp2.getRow(rowIndex);
			if (row != null) {
				Cell cell0 = row.getCell(0);
				Cell cell1 = row.getCell(1);
				if (cell0 != null && cell1 != null) {
					String current = cell0.getStringCellValue();
					String newCell = cell1.getStringCellValue();
					if (current != null && !current.trim().isEmpty()) {
						accordMap.put(current, newCell);
						System.out.println(current + " " + newCell);
					} else
						break;
				}
			}
		}
		System.out.println("最後一筆");
		
		String replacedString = "criteria";
		String title = replacedString + "修改紀錄";

//		String outFileName = "C:/Users/u10087/Desktop/criteriaModificationRecord.xls";
//		HSSFWorkbook workbook = new HSSFWorkbook();
//		FileOutputStream fileOut = null;
//		HSSFSheet sheet = workbook.createSheet(title);
//		sheet.setColumnWidth(0, 10000);
//		sheet.setColumnWidth(1, 10000);
//		HSSFRow rowTopic = sheet.createRow((short) 0); // rowCount ++
//		rowTopic.createCell(0).setCellValue(title);
//		HSSFRow rowhead = sheet.createRow((short) 1); // rowCount ++
//		rowhead.createCell(0).setCellValue("修改前");
//		rowhead.createCell(1).setCellValue("修改後");
//		int rowCount = 2;
		
		List<String> keyList = new ArrayList();
		for (Map.Entry<String, String> entry : accordMap.entrySet()) {
			String current = entry.getKey();
			//String newName = entry.getValue();
			keyList.add(current);
			}
		
		try {
			INode criteriaLib = myServer.getAdminInstance().getNode(NodeConstants.NODE_CRITERIA_LIBRARY);
			Object[] NodeList = criteriaLib.getChildren();
			for (Object object : NodeList) {

				INode nodeCriteria = (INode)object;
				String criteriaName = nodeCriteria.getName().toString();
				ICriteria criteria = (ICriteria) criteriaLib.getChild(criteriaName);

				// excel cell:criteria attributes
				String criteria_str = ConnecPLM.getCriteria_noNULL(criteria);
				for (String s : keyList) {
					if(criteria_str.toUpperCase().contains(s)) {
						System.out.println(criteriaName+"\n"+criteria_str);
					}
				}
				
				
				
//				for (Map.Entry<String, String> entry : accordMap.entrySet()) {
//					String current = entry.getKey();
//					String newName = entry.getValue();
//
//					INode node = (INode) object;
//					String criteriaName  = node.getName().toString();
//					String criteriaUpperName = node.getName().toString().toUpperCase();
//					
//					if (criteriaUpperName.contains(current)) {
//						
//						String  criteriaNewName = criteriaUpperName.replace(current, newName);
//						//node.setValue(NodeConstants.TYPE_CRITERIA,);
//						IProperty [] pList = node.getProperties();
//						for(IProperty p : pList) {
//							System.out.println(p.getId().toString()+p.getValue().toString());
//						
//							if(current =="PD"&&criteriaUpperName.startsWith("CPD")) {
//								continue;
//							}
//							else if(p.getValue().toString().equalsIgnoreCase(criteriaName)) {
//								p.setValue(criteriaNewName);
//								HSSFRow row = sheet.createRow(rowCount);
//								row.createCell(0).setCellValue(criteriaName);
//								row.createCell(1).setCellValue(criteriaNewName);
//								++rowCount;
//							}
//						}
//
//					}
//					
//
//				}

			}
		} catch (APIException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}

		
		// get a file output stream
//				try {
//					fileOut = new FileOutputStream(outFileName);
//				} catch (FileNotFoundException e) {
//					// TODO Auto-generated catch block
//					e.printStackTrace();
//				}
//
//				// write into excel
//				try {
//					workbook.write(fileOut);
//					fileOut.close();
//					System.out.println("Your excel file has been generated!");
//				} catch (IOException e) {
//					// TODO Auto-generated catch block
//					e.printStackTrace();
//				}
	}

}
