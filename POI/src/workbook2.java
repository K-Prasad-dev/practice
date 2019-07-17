import java.io.*;
import org.apache.poi.xssf.usermodel.*;

public class workbook2 {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		//Create blank workbook
		XSSFWorkbook workbook=new XSSFWorkbook();
		
		//Create file system using specific name
		FileOutputStream out= new FileOutputStream(new File("workbook2.xlsx"));
		
		//write operation workbook using file object
		workbook.write(out);
		out.close();
		System.out.println("workbook2.xlsx written successfully");
	}
}


