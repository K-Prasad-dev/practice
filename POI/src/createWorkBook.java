import java.io.*;
import org.apache.poi.xssf.usermodel.*;

public class createWorkBook {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		//Create blank workbook
		XSSFWorkbook workbook=new XSSFWorkbook();
		
		//Create file system using specific name
		FileOutputStream out= new FileOutputStream(new File("createworkbook.xlsx"));
		
		//write operation workbook using file object
		workbook.write(out);
		out.close();
		System.out.println("createworkbook.xlsx written successfully");
	}
}


