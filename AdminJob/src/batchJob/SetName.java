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

import com.agile.api.*;

import connect.ConnecPLM;

public class SetName {

	static void addStringinObj(IAgileSession myServer, Integer typeofObj, Integer whatToModify,
			List<String> objectsList, String addedStr, HSSFWorkbook workbook) {
		// IUserGroup.OBJECT_TYPE
		String title = addedStr + "修改紀錄";

		HSSFSheet sheet = workbook.createSheet(title);
		sheet.setColumnWidth(0, 8000);
		sheet.setColumnWidth(1, 8000);
		HSSFRow rowTopic = sheet.createRow((short) 0); // rowCount ++
		rowTopic.createCell(0).setCellValue(title);
		HSSFRow rowhead = sheet.createRow((short) 1); // rowCount ++
		rowhead.createCell(0).setCellValue("修改前");
		rowhead.createCell(1).setCellValue("修改後");
		int rowCount = 2;

		for (String originalName : objectsList) {
			try {
				if (originalName.startsWith("TW") || originalName.startsWith("JL") || originalName.startsWith("VL")) {
					if (myServer.getObject(typeofObj, originalName) != null) {
						IDataObject userGroup = (IDataObject) myServer.getObject(typeofObj, originalName);
						String newName = addedStr + originalName;
						userGroup.setValue(whatToModify, newName);
						HSSFRow row = sheet.createRow(rowCount);
						row.createCell(0).setCellValue(originalName);
						row.createCell(1).setCellValue(newName);
						++rowCount;
					} else {
						System.out.println("null");
					}
				}

			} catch (APIException e1) {
				e1.printStackTrace();
			}

		}

	}

	static void replaceStringinObj(IAgileSession myServer, Integer typeofObj, Integer whatToModify,
			List<String> objectsList, String replacedString, String replacingString, HSSFWorkbook workbook) {
		// IUserGroup.OBJECT_TYPE
		String title = replacedString + "修改紀錄";

		HSSFSheet sheet = workbook.createSheet(title);
		sheet.setColumnWidth(0, 8000);
		sheet.setColumnWidth(1, 8000);
		HSSFRow rowTopic = sheet.createRow((short) 0); // rowCount ++
		rowTopic.createCell(0).setCellValue(title);
		HSSFRow rowhead = sheet.createRow((short) 1); // rowCount ++
		rowhead.createCell(0).setCellValue("修改前");
		rowhead.createCell(1).setCellValue("修改後");
		int rowCount = 2;

		for (String originalName : objectsList) {
			try {
				if (originalName.contains(replacedString)) {
					if (myServer.getObject(typeofObj, originalName) != null) {
//						if(replacedString.equalsIgnoreCase("INACTIVE")) {
//							IDataObject userGroup= null;
//							String temp ="";
//							String newName = "";
//							
//							if(originalName.startsWith("INACTIVE ")) {
//								 userGroup = (IDataObject) myServer.getObject(typeofObj, originalName);
//								 temp = originalName.replaceAll(replacedString, "").trim();
//								 newName = "X_"+ temp;
//								 userGroup.setValue(whatToModify, newName);
//							}
//							else if(originalName.startsWith("INACTIVE_")) {
//								 userGroup = (IDataObject) myServer.getObject(typeofObj, originalName);
//								 temp = originalName.replaceAll(replacedString, "").trim();
//								 newName = "X"+ temp;
//								 userGroup.setValue(whatToModify, newName);
//								
//							}
//							else if(originalName.startsWith("INACTIVE")) {
//								 userGroup = (IDataObject) myServer.getObject(typeofObj, originalName);
//								 temp = originalName.replaceAll(replacedString, "").trim();
//								 newName = "X_"+ temp;
//								 userGroup.setValue(whatToModify, newName);
//								
//							}
//							
//							HSSFRow row = sheet.createRow(rowCount);
//							row.createCell(0).setCellValue(originalName);
//							row.createCell(1).setCellValue(newName);
//							++rowCount;
//						}else {
						IDataObject userGroup = (IDataObject) myServer.getObject(typeofObj, originalName);
						String newName = originalName.replaceAll(replacedString, replacingString);
						userGroup.setValue(whatToModify, newName);
						HSSFRow row = sheet.createRow(rowCount);
						row.createCell(0).setCellValue(originalName);
						row.createCell(1).setCellValue(newName);
						++rowCount;
//						}

					}
				}

			} catch (APIException e1) {
				e1.printStackTrace();
			}

		}

	}

	static int findFirstNumericStrIndex(String s) {
		for (int i = 0; i < s.length(); ++i) {
			char temp = s.charAt(i);
			if (temp <= 57 && temp >= 48) {
				return i;
			}
		}
		return -1;
	}

	static void replaceStringinUser(IAgileSession myServer, Integer typeofObj, Integer whatToModify,
			List<String> objectsList,Map accordMap, HSSFWorkbook workbook) {
		
		String title =  "user修改紀錄";

		HSSFSheet sheet = workbook.createSheet(title);
		sheet.setColumnWidth(0, 8000);
		sheet.setColumnWidth(1, 8000);
		HSSFRow rowTopic = sheet.createRow((short) 0); // rowCount ++
		rowTopic.createCell(0).setCellValue(title);
		HSSFRow rowhead = sheet.createRow((short) 1); // rowCount ++
		rowhead.createCell(0).setCellValue("修改前");
		rowhead.createCell(1).setCellValue("修改後");
		rowhead.createCell(2).setCellValue("部門");
		rowhead.createCell(3).setCellValue("email");
		
		int rowCount = 2;

		for (String originalName : objectsList) {
			try {

				if (myServer.getObject(typeofObj, originalName) != null) {
					System.out.println("get object");
					
					IDataObject user = (IDataObject) myServer.getObject(typeofObj, originalName);
					String FirstName = user.getValue(UserConstants.ATT_GENERAL_INFO_FIRST_NAME).toString();
					System.out.println(FirstName+ " " + originalName);
					int ID_index = findFirstNumericStrIndex(originalName);
					int FirstName_index = findFirstNumericStrIndex(FirstName);
					if(ID_index == -1 || FirstName_index== -1) {continue;}
					
					System.out.println(ID_index);
					String IDPrefix = originalName.substring(0, ID_index);
					String FirstNamePrefix = FirstName.substring(0, FirstName_index);

					if(accordMap.containsKey(FirstNamePrefix)) {
						String newName = accordMap.get(FirstNamePrefix) + originalName.substring(ID_index);
						System.out.println(newName);
						 user.setValue(whatToModify, newName);
						HSSFRow row = sheet.createRow(rowCount);
						row.createCell(0).setCellValue(originalName);
						row.createCell(1).setCellValue(newName);
						row.createCell(2).setCellValue(user.getValue(UserConstants.ATT_GENERAL_INFO_FIRST_NAME).toString());
						row.createCell(3).setCellValue(user.getValue(UserConstants.ATT_GENERAL_INFO_EMAIL).toString());
						++rowCount;
					}
					
				}

			} catch (APIException e1) {
				e1.printStackTrace();
			}

		}

	}

	public static void main(String[] args) {
		IAgileSession myServer = ConnecPLM.logInAsAdmin();
		List<String> userGroupList = new ArrayList<String>();
		Map<String, String> accordMap = new HashMap<>();

		// get a file input stream
		// allUserGroup.xlsx";
		String Infilename = "C:/Users/u10087/Desktop/PLM_INVALID_ID.xls";
		FileInputStream fis = null;
		HSSFWorkbook wb = null;
		try {
			fis = new FileInputStream(new File(Infilename));
			wb = new HSSFWorkbook(fis);
		} catch (FileNotFoundException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		} catch (IOException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		}

		HSSFSheet sheettemp = wb.getSheetAt(0);
		for (int rowIndex = 0; rowIndex <= sheettemp.getLastRowNum(); rowIndex++) {
			Row row = sheettemp.getRow(rowIndex);
			if (row != null) {
				Cell cell = row.getCell(0);
				if (cell != null) {
					String userGroup = cell.getStringCellValue();
					if (userGroup != null && !userGroup.trim().isEmpty()) {
						userGroupList.add(userGroup);
						System.out.println(userGroup);
					} else
						break;
				}
			}
		}
		System.out.println("最後一筆");

//		Integer type = IUserGroup.OBJECT_TYPE;
//		Integer toModify = UserGroupConstants.ATT_GENERAL_INFO_NAME;

		Integer type = IUser.OBJECT_TYPE;
		Integer toModify = UserConstants.ATT_GENERAL_INFO_USER_ID;

		String accordingToFile = "C:/Users/u10087/Desktop/IDreplaceList.xls";
		FileInputStream accordFS = null;
		HSSFWorkbook accordBook = null;
		try {
			accordFS = new FileInputStream(new File(accordingToFile));
			accordBook = new HSSFWorkbook(accordFS);
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

		String modifiedItems = "User";
		String outFileName = "C:/Users/u10087/Desktop/" + modifiedItems + "ModificationRecord.xls";
		HSSFWorkbook workbook = new HSSFWorkbook();
		FileOutputStream fileOut = null;

		replaceStringinUser(myServer, type, toModify, userGroupList,accordMap, workbook);
		
//		for (Map.Entry<String, String> entry : accordMap.entrySet()) {
//			String current = entry.getKey();
//			String newName = entry.getValue();
//			
//			replaceStringinObj(myServer, type, toModify, userGroupList, current, newName, workbook);
//		}

		// 要更新檔案才可以執行
		// addStringinObj(myServer, type, toModify, userGroupList,"UG_", workbook);

		// replaceStringinObj(myServer, type, toModify, userGroupList, "INACTIVE", "X_",
		// workbook);

		// get a file output stream
		try {
			fileOut = new FileOutputStream(outFileName);
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// write into excel
		try {
			workbook.write(fileOut);
			fileOut.close();
			System.out.println("Your excel file has been generated!");
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

}
