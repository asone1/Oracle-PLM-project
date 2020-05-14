package batchJob;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;

import com.agile.api.IAgileSession;

public class ConfirmADName {
	
	public static int find1stNon0Index(String s) {
		for (int i = 0; i < s.length(); ++i) {
			char temp = s.charAt(i);
			if (temp <= 57 && temp > 48) {
				return i;
			}
		}
		return -1;
	}

	public static int find1stNumericStrIndex(String s) {
		for (int i = 0; i < s.length(); ++i) {
			char temp = s.charAt(i);
			if (temp <= 57 && temp >= 48) {
				return i;
			}
		}
		return -1;
	}

	public static int count0(String s) {
		int count = 0;
		do {
			s = s.replaceFirst("0", "");
			++count;

		} while (s.contains("0"));

		return count;
	}

	public static boolean isInconsistent(String a, String b) {
		int indexA = find1stNumericStrIndex(a);
		int indexB = find1stNumericStrIndex(b);
		
		if (indexA == -1 || indexB == -1) {
			return false;
		} else {
			String AnumPart = a.substring(indexA);
			String BnumPart = b.substring(indexB);
			if (AnumPart.startsWith("0")) {
				int Non0indexA =find1stNon0Index(AnumPart);
				int Non0indexB =find1stNon0Index(BnumPart);
				if(Non0indexA==-1 || Non0indexB==-1) {return false;}
				String Anon0_numPart= AnumPart.substring(Non0indexA);
				String Bnon0_numPart= BnumPart.substring(Non0indexB);
				if(Anon0_numPart.equals(Bnon0_numPart)) {
					if (!a.substring(0, indexA).equalsIgnoreCase(b.substring(0, indexB))) {
						return true;
					}
					if (count0(a) != count0(b)) return true;
				}

			} else {
				if (AnumPart.equals(BnumPart)) {
					if (!a.substring(0, indexA).equalsIgnoreCase(b.substring(0, indexB))) {
						return true;
					}
				}
			}

			return false;
		}

	}

	public static void main(String[] args) {
		// IAgileSession myServer = ConnecPLM.logInAsAdmin();

		LinkedHashSet<String> PLMSet = new LinkedHashSet<String>();
		LinkedHashSet<String> ADSet = new LinkedHashSet<String>();

		// get a file input stream
		// allUserGroup.xlsx";
		String ADfilename = "C:/Users/u10087/Desktop/user_V1.xls";
		FileInputStream ADfis = null;
		HSSFWorkbook ADwb = null;
		try {
			ADfis = new FileInputStream(new File(ADfilename));
			ADwb = new HSSFWorkbook(ADfis);
		} catch (FileNotFoundException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		} catch (IOException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		}

		HSSFSheet sheetAD = ADwb.getSheetAt(0);
		for (int rowIndex = 0; rowIndex <= sheetAD.getLastRowNum(); rowIndex++) {
			Row row = sheetAD.getRow(rowIndex);
			if (row != null) {
				Cell cell = row.getCell(0);
				if (cell != null) {
					DataFormatter formatter = new DataFormatter();
					String userGroup = formatter.formatCellValue(cell);
					if (userGroup != null && !userGroup.trim().isEmpty()) {
						ADSet.add(userGroup);
					} else
						break;
				}
			}
		}
		System.out.println("AD最後一筆");

		// get a file input stream
		// allUserGroup.xlsx";
		String PLMfilename = "C:/Users/u10087/Desktop/plmActiveUser.xls";
		FileInputStream PLMfis = null;
		HSSFWorkbook PLMwb = null;
		try {
			PLMfis = new FileInputStream(new File(PLMfilename));
			PLMwb = new HSSFWorkbook(PLMfis);
		} catch (FileNotFoundException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		} catch (IOException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		}

		HSSFSheet sheetPLM = PLMwb.getSheetAt(0);
		for (int rowIndex = 0; rowIndex <= sheetPLM.getLastRowNum(); rowIndex++) {
			Row row = sheetPLM.getRow(rowIndex);
			if (row != null) {
				Cell cell = row.getCell(0);
				if (cell != null) {
					String userGroup = cell.getStringCellValue();
					if (userGroup != null && !userGroup.trim().isEmpty()) {
						PLMSet.add(userGroup);
					} else
						break;
				}
			}
		}
		System.out.println("PLM最後一筆");

		String title = "PLM與AD不一致";

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
//		rowhead.createCell(2).setCellValue("部門");
//		rowhead.createCell(3).setCellValue("email");

		for (String PLMId : PLMSet) {
			for (String ADId : ADSet) {
				if (isInconsistent(PLMId, ADId)) {
					HSSFRow row = sheet.createRow(rowCount);
					row.createCell(0).setCellValue(PLMId);
					row.createCell(1).setCellValue(ADId);
					++rowCount;
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
