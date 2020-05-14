package batchJob;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.NoSuchElementException;
import connect.*;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.agile.api.*;
import com.agile.px.ActionResult;
import com.agile.px.EventActionResult;


public class GetPartAndWhereUsed {

	HashSet<String> whereUsedFinalPartList = new HashSet<String>();
	static HashMap<String, String> dictionary = new HashMap<String, String>();

	public void findWhereUsed(IAgileSession myServer, String partNumber, Iterator It) {

		String temp = null;
		String component = "";
		boolean isEnd = true;
		try {
			for (int i = 100000; i > 0; --i) {

			//if(i == 9999) component = partNumber;
				
				while (It.hasNext()) {
					try {
						IRow nextRow = (IRow) (It.next());
						partNumber = nextRow.getValue(ItemConstants.ATT_WHERE_USED_ITEM_NUMBER).toString().trim();
						System.out.println(partNumber);
					} catch (NoSuchElementException e1) {
						break;
					}
					IItem part = (IItem) myServer.getObject(IItem.OBJECT_TYPE, partNumber);
					
					try {
						ITable tempTab = part.getTable(ItemConstants.TABLE_WHEREUSED);
						Iterator tempIt = tempTab.getTableIterator();
						IRow subRow = (IRow) (tempIt.next());
						temp = subRow.getValue(ItemConstants.ATT_WHERE_USED_ITEM_NUMBER).toString().trim();
					}
						catch (NoSuchElementException e1) {
						if (partNumber != null) {
							whereUsedFinalPartList.add(partNumber);
							//dictionary.put(component, temp);
							isEnd = false;
						}
					}
					catch (NullPointerException e1) {
						if (partNumber != null) {
							whereUsedFinalPartList.add(partNumber);
							System.out.println("add" +  partNumber+"null");
							isEnd = false;
						}
					}
					if (isEnd) {
						ITable subTab = part.getTable(ItemConstants.TABLE_WHEREUSED);
						Iterator subIt = subTab.getTableIterator();
						findWhereUsed(myServer, temp, subIt);
					}

					isEnd = true;

				}
			}
		}

		catch (APIException e1) {
			e1.printStackTrace();
		}

	}

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		IAgileSession myServer = ConnecPLM.logInAsAdmin();
		if (myServer != null) {
			System.out.println("connected");

			List<String> requiredPartList = new ArrayList<String>();

			// get a file input stream
			String Infilename = "C:/Users/u10087/Desktop/ICPSearchResults.xls";
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

			HSSFSheet sheet = wb.getSheetAt(0);
			for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
				Row row = sheet.getRow(rowIndex);
				if (row != null) {
					Cell cell = row.getCell(0);
					if (cell != null) {
						String partNum = cell.getStringCellValue();
						if (partNum != null && !partNum.trim().isEmpty()) {
							requiredPartList.add(partNum);
							System.out.println(partNum);
						} else
							break;
					}
				}
			}
			System.out.println("最後一筆");
			//requiredPartList.add("10G120N5JE04");
			

			String Outfilename = "C:/Users/u10087/Desktop/ICP_nonAVL_report.xls";
			FileOutputStream fileOut = null;
			HSSFWorkbook workbook = new HSSFWorkbook();
			HSSFSheet sheetOut = workbook.createSheet("0");
			sheetOut.setColumnWidth(0, 5000);
			sheetOut.setColumnWidth(1, 5000);
			HSSFRow rowTopic = sheetOut.createRow((short) 0); // rowCount ++
			rowTopic.createCell(0).setCellValue("料號BOM掛載零件規格書（系列），且指定非『AVL』的Manufacturer，且該料號申請單位為TW-4XXX的清單");
			HSSFRow rowhead = sheetOut.createRow((short) 1); // rowCount ++
			rowhead.createCell(0).setCellValue("符合條件料號");
			rowhead.createCell(1).setCellValue("引用此料的成品料");

			int count = 2;
			for (String part : requiredPartList) {
				try {
					System.out.println(part);
					HashMap params = new HashMap();
					params.put(ItemConstants.ATT_TITLE_BLOCK_NUMBER, part);
					IItem tempPart = (IItem) myServer.getObject(ItemConstants.CLASS_PARTS_CLASS, params);
					ITable tempTab = tempPart.getTable(ItemConstants.TABLE_WHEREUSED);
					Iterator tempIt = tempTab.getTableIterator();
					GetPartAndWhereUsed requiredPart = new GetPartAndWhereUsed();
					requiredPart.findWhereUsed(myServer, part, tempIt);
					String topUsedPartString = "";
					
					for (String temp : requiredPart.whereUsedFinalPartList) {
						topUsedPartString += temp + "；";
					}
					
					
					if(requiredPart.whereUsedFinalPartList.isEmpty()) {
						HSSFRow newRow = sheetOut.createRow((short) count);
						newRow.createCell(0).setCellValue(part);
						newRow.createCell(1).setCellValue("whereUsed為空");
						count++;
					}else {
						HSSFRow rowMain = sheetOut.createRow((short) count);
						rowMain.createCell(0).setCellValue(part);
						rowMain.createCell(1).setCellValue(topUsedPartString);
						count++;
					}
				} catch (NoSuchElementException e1) {
					HSSFRow rowMain = sheetOut.createRow((short) count);
					rowMain.createCell(0).setCellValue(part);
					rowMain.createCell(1).setCellValue("whereUsed為空(NoSuchElement)");
					count++;
				} catch (NullPointerException e1) {
					HSSFRow rowMain = sheetOut.createRow((short) count);
					rowMain.createCell(0).setCellValue(part);
					rowMain.createCell(1).setCellValue("whereUsed為空(NullPointer)");
					count++;
				} catch (APIException e1) {
					e1.printStackTrace();
				}
			}

			try {
				fileOut = new FileOutputStream(Outfilename);
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

			System.out.println("END");
		} else {
			System.out.println("Connection is null");
		}

	}

}
