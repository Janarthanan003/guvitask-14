package writeexcel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {
	public static void main(String []args) throws FileNotFoundException, IOException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Sheet1");
		Object [][] data= {
				{"Name","Age","E-mail"},
				{"John Doe",30,"john@test.com"},		
				{"Jane Doe",28,"john@test.com"},
				{"Bob Smith",35,"jacky@example.com"},
				{"Swapnil",37,"swapnil@example.com"}
		};
		int rowNum=0;
		for (Object[] rowdata : data) {
			XSSFRow row=sheet.createRow(rowNum++);
			int colNum=0;
			for (Object field: rowdata) {
				XSSFCell cell=row.createCell(colNum++);
				if (field instanceof String) {
					cell.setCellValue((String)field);
				}else if(field instanceof Integer) {
					cell.setCellValue((Integer) field);
				}
			}
		}
		try(FileOutputStream outputstream= new FileOutputStream("data.xlsx"))
		{
			workbook.write(outputstream);
		}			
		System.out.println("Data successfully written to file");
	} 
}