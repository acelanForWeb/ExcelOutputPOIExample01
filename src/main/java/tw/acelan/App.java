package tw.acelan;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class App {
	//excel檔案輸出的路徑，可以自行喜好設定，請確認路徑相關資料夾已存在
	private static final String FILE_NAME = "D:/Demo/PhoneBook.xlsx";
	
	public static void main(String[] args){
		//建立一個excel work book物件，等同於建立一個Excel檔案
		XSSFWorkbook workbook = new XSSFWorkbook();
		System.out.println("建立excel檔案完成");
		
		
		//在excel work book物件中建立一個「sheet」，描述為「電話簿」
		XSSFSheet sheet = workbook.createSheet("電話簿");
		System.out.println("建立sheet完成");
		
		//電話簿資料
		Object[][] phoneBook = {
			{"姓名", "生日", "手機號碼"},	
			{"A君", "1970/01/25", "0900-111-111"},
            {"B君", "1980/02/26", "0900-222-222"},
            {"C君", "1990/05/25", "0900-333-333"},
            {"D君", "2000/07/13", "0900-444-444"},
            {"E君", "2010/11/21", "0900-555-555"}
		};
		
		
		//====		將電話簿資料填充到sheet中		Begin		====//
		System.out.println("開始將資料寫入sheet...");
		int rowNum = 0;
		for (Object[] rowData : phoneBook) {
			 Row row = sheet.createRow(rowNum++);
			 int colNum = 0;
			 for (Object field : rowData) {
				 Cell cell = row.createCell(colNum++);
				 if (field instanceof String) {
					 cell.setCellValue((String) field);
				 }else if(field instanceof Integer){
					 cell.setCellValue((Integer) field);
				 }
			 }
		}
		System.out.println("資料寫入sheet完成");
		//====		將電話簿資料填充到sheet中		End			====//
		
		
		//====		輸出Excel檔案				Begin		====//
		System.out.println("開始將excel檔案進行輸出...");
		FileOutputStream outputStream = null;
		try {
            outputStream = new FileOutputStream(FILE_NAME);
            workbook.write(outputStream);
           
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }finally{
        	try{
        		if(workbook != null){
        			workbook.close();
        		}
        	}catch(Exception e){e.printStackTrace();}
        	
        	try{
        		if(outputStream != null){
        			outputStream.close();
        		}
        	}catch(Exception e){e.printStackTrace();}
        }
		System.out.println("excel檔案輸出完成");
		//====		輸出Excel檔案				End			====//
	}
}
