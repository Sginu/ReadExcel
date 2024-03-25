
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadDignifyData {

	
	private static String filePath = "C:\\dignify_data.xlsx";
	
	public static void main(String[] args) {
			readTheFile();
		
	}
	
	private static void readTheFile() {
		 try {
		     Workbook workbook =  new XSSFWorkbook("C:\\data\\test.xlsx");
	         int numberOfSheets = workbook.getNumberOfSheets();
	         String dlrCd = null;
	         String url = null;
	         int cnt = 0; int c = 0;
	         for (int i = 0; i < numberOfSheets; i++) {
	                Sheet sheet = workbook.getSheetAt(i);
	                Iterator rowIterator = sheet.iterator();
	                
	              //iterating over each row
	                while (rowIterator.hasNext()) {
						cnt++;
	                	Row row = (Row) rowIterator.next();
	                    Iterator cellIterator = row.cellIterator();
	                    
	                    while (cellIterator.hasNext()) {
	                    	
	                    	Cell cell = (Cell) cellIterator.next();
	                    	 
	                    	if (cell.getColumnIndex() == 0) {
	                    		dlrCd = String.valueOf(cell.getStringCellValue());
                            }
	                    	//Cell with index 2 contains marks in Science
                            else if (cell.getColumnIndex() == 1) {
                            	url = String.valueOf(cell.getStringCellValue());
                            }
	                    }
						System.out.println("ELSIF (DLR.DLR_CD = "+dlrCd+") THEN DLR_URL := '"+url+"';");
	                	
	                }
	                System.out.println("No of Rows :"+cnt);
	                
	         }
		 } catch (FileNotFoundException e) {
	            e.printStackTrace();
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	}

}
