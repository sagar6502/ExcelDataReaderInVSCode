package com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestDataReader {
    public  void readData(String dataFile1) throws IOException {
		FileInputStream fis = new FileInputStream(new File(dataFile1));
		XSSFWorkbook workbook = new XSSFWorkbook(fis); 
		
		Sheet sheet = workbook.getSheet("Test");
        readColumnNames(dataFile1, "Configuration");
		
		for(Row row : sheet) {
			for(Cell cell : row) {
				String cellValue = cell.getStringCellValue();
				System.out.println(cellValue +"\t");
			}
			System.out.println();
		}
		workbook.close();
	}

    public void writeExcelFile(String dataFile) throws IOException {
        File file = new File(dataFile);
        String columnName = "Module3";
        String valueToUpdate = "Reset2";

        if (file.exists() && file.length() > 0) {
            try (FileInputStream fis = new FileInputStream(file)) {
                XSSFWorkbook workbook = new XSSFWorkbook(fis);
                Sheet sheet = workbook.getSheet("Test");

                int dataColumnIndex = getColumnIndex(dataFile, columnName); // Assuming you want data from the second column (index 1)

                for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);

                    if (row == null) {
                        System.out.println("Skipping null row at index " + i);
                        continue;
                    }

                    Cell dataCell = row.getCell(dataColumnIndex);
                    if (dataCell == null) {
                        dataCell = row.createCell(dataColumnIndex);
                    }
                    if(!isValueExists(dataFile, columnName,valueToUpdate)){
                        dataCell.setCellValue(valueToUpdate);
                        try (FileOutputStream fos = new FileOutputStream(file)) {
                          workbook.write(fos);
                            System.out.println("Workbook written successfully.");
                        }
                    }
                        
                }

                // Write the modified workbook back to the same file
                
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            System.err.println("The file is empty or does not exist.");
        }
    }
    public int getColumnIndex(String dataFile, String columnName) throws IOException {
        File file = new File(dataFile);

        if (file.exists() && file.length() > 0) {
            try (FileInputStream fis = new FileInputStream(file)) {
                XSSFWorkbook workbook = new XSSFWorkbook(fis);
                Sheet sheet = workbook.getSheet("Test");

                // Get the first row (header row) to find the column index
                Row headerRow = sheet.getRow(0);

                if (headerRow != null) {
                    // Iterate through cells in the header row to find the matching column name
                    for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                        Cell cell = headerRow.getCell(i);

                        if (cell != null && cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                            return i;
                        }
                    }
                }

                // Column not found
                System.err.println("Column '" + columnName + "' not found.");
                return -1;
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            System.err.println("The file is empty or does not exist.");
        }

        return -1;
    }
    /* 
    private static int getTotalColumns(String dataFile) throws IOException {
        File file = new File(dataFile);

        if (file.exists() && file.length() > 0) {
            System.out.println("File exists");

            try (FileInputStream fis = new FileInputStream(file)) {
                XSSFWorkbook workbook = new XSSFWorkbook(fis);
                Sheet sheet = workbook.getSheet("Test");

                Row headerRow = sheet.getRow(0);

                if (headerRow != null) {
                    return headerRow.getLastCellNum();
                } else {
                    // No header row found
                    System.err.println("No header row found in the sheet.");
                    return 0;
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            System.err.println("The file is empty or does not exist.");
        }

        return 0;
    }
    */

    private static boolean isValueExists(String dataFile, String columnName, String valueToCheck) throws IOException {
        File file = new File(dataFile);

        if (file.exists() && file.length() > 0) {
            
            try (FileInputStream fis = new FileInputStream(file)) {
                XSSFWorkbook workbook = new XSSFWorkbook(fis);
                Sheet sheet = workbook.getSheet("Test");

                // Get the column index based on the column name
                int columnIndex = getColumnIndex(sheet, columnName);

                if (columnIndex != -1) {
                    // Iterate through rows to check if the value exists in the specified column
                    for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                        Row row = sheet.getRow(i);

                        if (row != null) {
                            Cell cell = row.getCell(columnIndex);

                            if (cell != null && cell.getStringCellValue().equalsIgnoreCase(valueToCheck)) {
                                System.out.println("Duplicate value found of :: "+valueToCheck+ " at row no :: "+i);
                                return true; // Value found
                            }
                        }
                    }
                } else {
                    System.err.println("Column '" + columnName + "' not found.");
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            System.err.println("The file is empty or does not exist.");
        }

        return false; // Value not found or encountered an error
    }

    private static int getColumnIndex(Sheet sheet, String columnName) {
        Row headerRow = sheet.getRow(0);

        if (headerRow != null) {
            // Iterate through cells in the header row to find the matching column name
            for (int i = 0; i < headerRow.getLastCellNum(); i++) {
                Cell cell = headerRow.getCell(i);

                if (cell != null && cell.getStringCellValue().equalsIgnoreCase(columnName)) {
                    return i; // Column index found
                }
            }
        }

        return -1; // Column not found
    }

    public  void readColumnNames(String dataFile, String sheetName) throws IOException {
		FileInputStream fis = new FileInputStream(new File(dataFile));
		XSSFWorkbook workbook = new XSSFWorkbook(fis); 
		
		Sheet sheet = workbook.getSheet(sheetName);
		Row row = sheet.getRow(0);
		System.out.println("Column Names are :: ");
			for(Cell cell : row) {
				String cellValue = cell.getStringCellValue();
				System.out.print(cellValue +" ");
			}
			//System.out.println();
		workbook.close();
	}
    
}
