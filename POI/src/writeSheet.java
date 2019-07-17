import java.io.*;
import java.io.FileOutputStream;

import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.*;
public class writeSheet {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		File file=new File("createWorkBook.xlsx");
		FileInputStream fIP= new FileInputStream(file);
		
		//Get the workbook instance for XLSX file
		XSSFWorkbook workbook= new XSSFWorkbook(fIP);
		
		if(file.isFile()&&file.exists()) {
			System.out.println("createWorkBook.xlsx file open successfully.");
		}
		else {
			System.out.println("Error to open createWorkBook.xlsx file.");
		}
		 XSSFSheet spreadsheet = workbook.getSheet("Student Data");
		//Create row object
		//XSSFRow row;
		
		//create a row in the sheet
		XSSFRow row=spreadsheet.createRow(0);
		XSSFRow row1=spreadsheet.createRow(1);
		
		//create cell in the sheet
		XSSFCell cell1= row.createCell(0);
		
		//This data needs to be written (Object[])
		Map <String, Object[] > stinfo=
		new TreeMap < String, Object[] >();
		stinfo.put("1",  new Object[] { "ST ID", "ST NAME", "COURSE"});
		stinfo.put("2", new Object[] { "1529010064", "Khushi Prasad", "B.tech, CSE"});
		stinfo.put("3", new Object[] {"1529010076", "Md. Sajid Khan", "B.tech, CSE"});
		stinfo.put("4", new Object[] {"1429013036", "Vaibhav Sadhna", "B.tech, IT"});
		
		//Iterate over data and write to sheet
		Set < String > keyid= stinfo.keySet();
		int rowid=0;
		System.out.println(spreadsheet);
		for(String key: keyid) {
			row=spreadsheet.createRow(rowid++);
			Object [] objectArr = stinfo.get(key);
			int cellid= 0;
			
			for(Object obj : objectArr) {
				Cell cell= row.createCell(cellid++);
				cell.setCellValue((String)obj);
			}
		}
		
		//Write the workbook in file system
		FileOutputStream out= new FileOutputStream("createWorkBook.xlsx");
		workbook.write(out);
		out.close();
		System.out.println("writeSheet.xlsx written successfully");
			}
			
		

	}


