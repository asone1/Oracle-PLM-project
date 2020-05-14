package batchJob;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import connect.*;
import com.agile.api.*;

public class CancelChange {
	static final String user = "t35756";

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		IAgileSession myServer = ConnecPLM.logInAsAdmin();
		List<String> toBeCanceled = new ArrayList<String>();
		// toBeCanceled.add("PR102159");

		String Infilename = "C:/Users/u10087/Desktop/tobeCanceled.xls";
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
						toBeCanceled.add(userGroup);
						System.out.println(userGroup);
					} else
						break;
				}
			}
		}
		System.out.println("最後一筆");

		String outFileName = "C:/Users/u10087/Desktop/" + user + "ModificationRecord.xls";
		HSSFWorkbook workbook = new HSSFWorkbook();
		FileOutputStream fileOut = null;
		String title = user + "表單修改紀錄";

		HSSFSheet sheet = workbook.createSheet(title);
		sheet.setColumnWidth(0, 8000);
		sheet.setColumnWidth(1, 8000);
		HSSFRow rowTopic = sheet.createRow((short) 0); // rowCount ++
		rowTopic.createCell(0).setCellValue(title);
		HSSFRow rowhead = sheet.createRow((short) 1); // rowCount ++
		rowhead.createCell(0).setCellValue("修改前");
		rowhead.createCell(1).setCellValue("修改後");
		int rowCount = 2;

		for (String change : toBeCanceled) {
			try {
				IChange changeObj = (IChange) myServer.getObject(IChange.OBJECT_TYPE, change);
				if (changeObj.getValue(ChangeConstants.ATT_COVER_PAGE_ORIGINATOR).toString().contains(user)) {

					ITable affectedTab = changeObj.getTable(ChangeConstants.TABLE_AFFECTEDITEMS);
					@SuppressWarnings("rawtypes")
					Iterator affectedIter = affectedTab.iterator();
					String newData = "";
					HSSFRow excelRow = sheet.createRow(rowCount);
					
					while (affectedIter.hasNext()) {
						IRow affectedRow = (IRow) (affectedIter.next());
						IItem part = (IItem) affectedRow.getReferent();
						ITable redlineBOM = part.getTable(ItemConstants.TABLE_REDLINEBOM);
						Iterator it = redlineBOM.iterator();
						while (it.hasNext()) {
							IRedlinedRow row = (IRedlinedRow) it.next();
							row.undoRedline();
							newData += "undoRedline/";
						}
					}

					changeObj.refresh();

					while (affectedIter.hasNext()) {
						IRow affectedRow = (IRow) (affectedIter.next());
						IItem part = (IItem) affectedRow.getReferent();
						ITable attachment = part.getTable(ItemConstants.TABLE_ATTACHMENTS);
						Iterator attchIter = attachment.iterator();
						while (attchIter.hasNext()) {
							IRow row = (IRow) attchIter.next();
							attachment.removeRow(row);
							newData += "remove attachment";
						}
					}

					changeObj.refresh();
					IStatus[] nextStatuses = changeObj.getNextStatuses();
					IStatus Canceled = null;
					for (int i = 0; i < nextStatuses.length; i++) {
						System.out.println("nextStatuses[" + nextStatuses[i].getName() + "] = ");
						if (nextStatuses[i].getName().toString().equalsIgnoreCase("Canceled")) {
							Canceled = nextStatuses[i];
						}
					}
					
					//若站點為unassigned
//					if(Canceled == null) {
//						
//						IStatus initiate = null;
//						for (int i = 0; i < nextStatuses.length; i++) {
//							System.out.println("nextStatuses[" + nextStatuses[i].getName() + "] = ");
//							if (nextStatuses[i].getName().toString().equalsIgnoreCase("Initiate")) {
//								Canceled = nextStatuses[i];
//							}
//						}
//					}

					Collection userList = new ArrayList();
					userList.add(myServer.getObject(IUser.OBJECT_TYPE, "plmadmin"));
					// changeObj.getReviewers

					excelRow.createCell(0).setCellValue(change);
					excelRow.createCell(1).setCellValue(newData);
					++rowCount;

					changeObj.changeStatus(Canceled, false, "根據嘉隆資訊需求申請單：JL200400657", false, false, userList, userList,
							null, null, false);

				}

			} catch (APIException ae) {
				ae.printStackTrace();

				System.out.println("END");
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
