package practice;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Arrays;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel_file_reading {
	 Object[][] abc = null;
	public void information() {
		FileInputStream file = null;
		try {
			file = new FileInputStream("D:\\Selinium practice\\keyward_driven\\Book1.xlsx");

			try {
				XSSFWorkbook book = new XSSFWorkbook(file);
				XSSFSheet sheet = book.getSheet("Sheet1");
				int row = sheet.getLastRowNum();
				XSSFRow row1 = sheet.getRow(0);
				 int column=row1.getLastCellNum();
				  abc=new Object[row+1][column];
				 System.out.println("["+abc.length+"]"+"["+abc[row].length+"]");
				 System.out.println("row : "+ row + " column : "+ column);
				for (int i=0;i<=row;i++)
				{
					row1= sheet.getRow(i);
					for(int j=0;j<column;j++) {
						abc[i][j]=row1.getCell(j);
						System.out.println( abc[i][j]);
					}
				}
				

			} catch (FileNotFoundException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	public static void main(String[] args) {
		Excel_file_reading obj = new Excel_file_reading();
		obj.information();
		
	}

}
