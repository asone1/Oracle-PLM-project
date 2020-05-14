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

public class setUsetName {

	public static void main(String[] args) {
		IAgileSession myServer = ConnecPLM.logInAsAdmin();

		// get a file input stream
		// allUserGroup.xlsx";
		String Infilename = "C:/Users/u10087/Desktop/userToBe.xls";
		FileInputStream fis = null;
		HSSFWorkbook wb = null;
		try {
			fis = new FileInputStream(new File(Infilename));
			wb = new HSSFWorkbook(fis);
			System.out.println("IN");
		} catch (FileNotFoundException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		} catch (IOException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		}

		String title = "PLM­×§ï¬ö¿ý";
		String modifiedItems = "UserID";
		String outFileName = "C:/Users/u10087/Desktop/" + modifiedItems + "ModificationRecord.xls";
		HSSFWorkbook workbook = new HSSFWorkbook();
		FileOutputStream fileOut = null;
		HSSFSheet sheet = workbook.createSheet(title);
		int rowCount = 2;
		sheet.setColumnWidth(0, 8000);
		sheet.setColumnWidth(1, 8000);
		HSSFRow rowTopic = sheet.createRow((short) 0); // rowCount ++
		rowTopic.createCell(0).setCellValue(title);
		HSSFRow rowhead = sheet.createRow((short) 1); // rowCount ++
		rowhead.createCell(0).setCellValue("PLM");
		rowhead.createCell(1).setCellValue("AD");

		HSSFSheet sheettemp = wb.getSheetAt(0);
		for (int rowIndex = 0; rowIndex <= sheettemp.getLastRowNum(); rowIndex++) {
			Row row = sheettemp.getRow(rowIndex);
			if (row != null) {
				System.out.println("1");
				Cell cell1 = row.getCell(0);
				Cell cell2 = row.getCell(1);
				Cell cell3 = row.getCell(2);
				if (cell1 != null && cell2 != null && cell3 != null) {
					String userPLM = cell1.getStringCellValue();
					// String userName = cell2.getStringCellValue();
					String ADID = cell3.getStringCellValue();
					Integer toModify = UserConstants.ATT_GENERAL_INFO_USER_ID;
					System.out.println("2");
					try {
						if (myServer.getObject(IUser.OBJECT_TYPE, userPLM) != null) {
							IUser user = (IUser) myServer.getObject(IUser.OBJECT_TYPE, userPLM);
							user.setValue(UserConstants.ATT_GENERAL_INFO_USER_ID, ADID);
							System.out.println("set"+userPLM);
							HSSFRow temp = sheet.createRow(rowCount);
							temp.createCell(0).setCellValue(userPLM);
							temp.createCell(1).setCellValue(ADID);
							++rowCount;
						} else {
							System.out.println("null");
						}
					} catch (APIException e1) {
						e1.printStackTrace();
					}

				}
			}
		}

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
