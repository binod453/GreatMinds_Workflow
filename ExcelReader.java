package practice;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {
	public static void main(String[] args) throws IOException, InvalidFormatException {
		String excelFilePath = "C:\\Users\\binod.moharana\\Desktop\\Bynder Workflow SRS\\ExceltoJson\\greatminds_wf_2019-01-16.xlsx";
		DataFormatter dataFormatter = new DataFormatter();
		try {
			FileInputStream fInputStream = new FileInputStream(excelFilePath.trim());
			Workbook excelWorkBook = new XSSFWorkbook(fInputStream);
			int totalSheetNumber = excelWorkBook.getNumberOfSheets();
			// for (int i = 0; i < totalSheetNumber; i++) {
			Sheet sheet = excelWorkBook.getSheetAt(0);
			String sheetName = sheet.getSheetName();
			Row row = null;
			Map<String, Integer> duplicateJobCountStart = new HashMap<>();
			Map<String, Integer> duplicateJobCountEnd = new HashMap<>();
			if (sheetName != null && sheetName.length() > 0) {
				for (int rowIndex = 1; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
					row = sheet.getRow(rowIndex);
					if (row != null) {
						Cell cell = row.getCell(2);
						if (duplicateJobCountStart != null && !duplicateJobCountStart.containsKey(cell.toString()))
							duplicateJobCountStart.put(cell.toString(), rowIndex);
						if (duplicateJobCountEnd != null && !duplicateJobCountEnd.containsKey(cell))
							duplicateJobCountEnd.put(cell.toString(), rowIndex);
					}
				}
			}
			int count = 17;
			Map<Integer, List<String>> newData = new HashMap<Integer, List<String>>();
			// List<String> dataList = new ArrayList<String>();
			for (Entry<String, Integer> entry : duplicateJobCountStart.entrySet()) {
				for (Entry<String, Integer> entry2 : duplicateJobCountEnd.entrySet()) {
					if (entry.getKey().equalsIgnoreCase(entry2.getKey())) {
						System.out.println(entry.getKey() + "====" + entry.getValue() + "====" + entry2.getValue());
						List<String> dataList = new ArrayList<String>();
						for (int j = entry.getValue(); j <= entry2.getValue(); j++) {
							Row firstRow = sheet.getRow(entry.getValue());
							// int cellOnRow = entry.getValue();
							row = sheet.getRow(j);
							if (row != null) {
								String stage_name = dataFormatter.formatCellValue(row.getCell(11));
								String stage_position = dataFormatter.formatCellValue(row.getCell(12));
								String stage_status = dataFormatter.formatCellValue(row.getCell(13));
								String stage_responsible = dataFormatter.formatCellValue(row.getCell(14));
								String stage_date_started = dataFormatter.formatCellValue(row.getCell(15));
								String stage_date_finished = dataFormatter.formatCellValue(row.getCell(16));
								String stage_duration = dataFormatter.formatCellValue(row.getCell(17));

								dataList.add(stage_name);
								dataList.add(stage_position);
								dataList.add(stage_status);
								dataList.add(stage_responsible);
								dataList.add(stage_date_started);
								dataList.add(stage_date_finished);
								dataList.add(stage_duration);

							}
						}
						newData.put(entry.getValue(), dataList);
						Set<Integer> newRows = newData.keySet();
						// int rownum = ;

						for (Integer key : newRows) {
							Row newRow = sheet.getRow(key);
							List<String> objArr = newData.get(key);
							int cellnum = 18;
							for (Object obj : objArr) {
								Cell cell = newRow.createCell(cellnum++);
								if (obj instanceof String) {
									cell.setCellValue((String) obj);
								} else if (obj instanceof Boolean) {
									cell.setCellValue((Boolean) obj);
								} else if (obj instanceof Date) {
									cell.setCellValue((Date) obj);
								} else if (obj instanceof Double) {
									cell.setCellValue((Double) obj);
								}
							}
						}

						// read row wise columns and insert accordingly
					}
				}
			}
			// }

			// readexcel(sheet, dataFormatter);

			FileOutputStream os = new FileOutputStream(new File(excelFilePath));
			excelWorkBook.write(os);
			System.out.println("Writing on Excel file Finished ...");

			// Close workbook, OutputStream and Excel file to prevent leak
			os.close();

			excelWorkBook.close();
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	private static void readexcel(Sheet sheet, DataFormatter dataFormatter) {

		System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
		sheet.forEach(row -> {
			row.forEach(cell -> {
				String cellValue = dataFormatter.formatCellValue(cell);
				System.out.print(cellValue + "\t");
			});
			System.out.println();
		});
	}
}
