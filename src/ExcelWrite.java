import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelWrite {
public static void main(String[] args) throws EncryptedDocumentException, FileNotFoundException, IOException {
		
		String path1 = "./data/data.xlsx";
		Workbook wb = WorkbookFactory.create(new FileInputStream(path1));
		
		wb.getSheet("Sheet1").getRow(1).getCell(0).setCellValue("admin");
		
		//To save 
		wb.write(new FileOutputStream(path));

  }
}
