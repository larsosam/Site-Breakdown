import  java.io.*;
import java.util.Arrays;
import java.util.Scanner;

import com.monitorjbl.xlsx.StreamingReader;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class mainClass {

	public static void main(String[] args) {
		
		//Makes sure we have valid arguments
		if (args.length != 2) {
			System.out.println("Usage: java -jar SiteBreakdown.jar [Award Database Path] [Output Path]");
			System.exit(1);
		}
		
		String databasePath = args[0];
		if (databasePath.endsWith(File.separator)) {
			databasePath = databasePath.substring(0, databasePath.length());
		}
		String wbPath = args[1];
		if (!wbPath.endsWith(File.separator)) {
			wbPath += File.separator;
		}
		
		//Incorrect type for database
		if (!databasePath.endsWith("xlsx")) {
			System.out.println("Incorrect file extension for award database. Please use .xlsx");
			System.exit(1);
		}
		
		//Checks to see if the destination path is a directory
		if (!(new File(wbPath).isDirectory())) {
			System.out.println("Destination path either does not, or is not a directory.");
			System.exit(1);
		}
		
		//If output file exists, looks if you want to overwrite
		if (new File(wbPath + "Site Based Database.xlsx").exists() ||
				new File(wbPath + "Site Based Database.xlsx").exists()) {
			System.out.println("Output file already exists. Overwrite? y/n");
			boolean hasChoice = false;
			while (hasChoice == false) {
				Scanner scnr = new Scanner(System.in);
				String choice = scnr.next();
				if (choice.equals("y")) {
					hasChoice = true;
					scnr.close();
				} else if (choice.equals("n")) {
					System.out.println("Exiting Program");
					System.exit(0);
				} else {
					System.out.println("Invalid Input. y for yes, n for no");
				}
			}

		}
		
		Workbook workbook = null;
		InputStream is = null;
		
		//Creates the workbook stream
		try {
			is = new FileInputStream(new File(databasePath));
			workbook = StreamingReader.builder()
			        .rowCacheSize(100)    // number of rows to keep in memory (defaults to 10)
			        .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
			        .open(is);            // InputStream or File for XLSX file (required)
	
		} catch (Exception e) {
			System.out.println("Database either could not be opened, or could not be found");
			System.exit(1);
		}
		
	
		System.out.println("Database opened");

		//The institutions we want to look at
		String[] desiredInstitutions = {""};
		
		//The workbooks and stuff to write to
		SXSSFWorkbook siteBased = new SXSSFWorkbook(10);
		SXSSFSheet siteSheet = siteBased.createSheet();
		SXSSFRow siteHeader = siteSheet.createRow(0);
		
		//Goes through every row, copies values if meets the requirements
		for (Row row : workbook.getSheetAt(0)) {
			try {

				//For the header row
				if (row.getCell(0).getRowIndex() == 0) {
					for (Cell cell : row) {
						
						//Starts out the row at column 0, because lastCellNum is weird
						if (cell.getColumnIndex() == 0) {
							siteHeader.createCell(0).setCellValue(cell.getStringCellValue());
						}
						else {
							siteHeader.createCell(siteHeader.getLastCellNum()).
							setCellValue(cell.getStringCellValue());
						}
					}
				}
				
				//For all other rows, if fits in date range, 
				//and from a good institution
				else if (Integer.parseInt(row.getCell(1).getStringCellValue().substring(6, 10)) >= 1995 && 
						Arrays.asList(desiredInstitutions).contains(row.getCell(41).getStringCellValue())) {
					siteSheet.createRow(siteSheet.getLastRowNum() + 1);
					for (int i = 0; i < row.getLastCellNum(); i++) {
						try {
							siteSheet.getRow(siteSheet.getLastRowNum()).
							createCell(i).setCellValue(row.getCell(i).getStringCellValue());	
						} 
						//Just in case something weird happens
						catch (NullPointerException ex) {
							siteSheet.getRow(siteSheet.getLastRowNum()).createCell(i).setCellValue("");
						}
					}
				}	
			//In case the formatting is off
			} catch (NumberFormatException ex) {
				System.out.println("ERROR WITH: " + row.getCell(1).getStringCellValue());
			}
		}
		System.out.println("Writing Output");
		//Write the output
		try {
			FileOutputStream fileOut = new FileOutputStream(wbPath + "Site Based Database.xlsx");
			siteBased.write(fileOut);
			fileOut.close();
			System.out.println("File successfully written to " +wbPath + "Site Based Database.xlsx");
			siteBased.close();
		} catch (IOException e) {
			System.out.println("Somehow the file you want to write to no longer exists. Cmon.");
		}
	}
}
