import java.io.*;
import org.apache.poi.hssf.usermodel.*;

public class exampleOpen {

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub
		File file= new File("D:\\example.xls");
		FileInputStream fIP= new FileInputStream(file);
		
		//Get the workbook instance for XLSX file
		HSSFWorkbook workbook= new HSSFWorkbook(fIP);
		
		if(file.isFile()&& file.exists()) {
			System.out.println("exampleOpen.xlsx file open successfully.");
		}
		else {
			System.out.println("Error to open exampleOpen.xlsx file.");
		}
		}
	}


