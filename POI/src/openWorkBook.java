import java.io.*;
import org.apache.poi.xssf.usermodel.*;

public class openWorkBook {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		File file= new File("example.xlsx");
		FileInputStream fIP= new FileInputStream(file);
		
		//Get the workbook instance for XLSX file
		XSSFWorkbook workbook= new XSSFWorkbook(fIP);
		
		if(file.isFile()&& file.exists()) {
			System.out.println("createworkbook.xlsx file open successfully.");
		}
		else {
			System.out.println("Error to open openworkbook.xlsx file.");
		}
		}

	}


