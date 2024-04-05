package testapp1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoiDemo {

	public static void main(String[] args) {
	//	createworkbook("employees","records");
		readexcel();
		//appendrow("employees", "records", "5", "admin", "Lilo");
	}
	
	public static void appendrow(String wbook, String wsheet, String id, String dept, String name) {
		
		try {
		
			FileInputStream file = new FileInputStream(new File(wbook + ".xlsx"));
			
			
				
				XSSFWorkbook workbook = new XSSFWorkbook(file);
				XSSFSheet sheet = workbook.getSheet(wsheet);
				
				//to get the last row record
				
				int rowlastnum = sheet.getLastRowNum();
				Row newrow = sheet.createRow(rowlastnum + 1);
				
				//newrow.createCell(0).setCellValue(id)); same sa baba
				
				Cell cell1 = newrow.createCell(0);
				cell1.setCellValue(id);
				
				Cell cell2 = newrow.createCell(1);
				cell2.setCellValue(name);
				
				Cell cell3 = newrow.createCell(2);
				cell3.setCellValue(dept);
				
				//write to file
				
				FileOutputStream out = new FileOutputStream(new File(wbook + ".xlsx"));
				workbook.write(out);
				System.out.println("new row added");
				
				out.close();
				
			
		} catch (Exception e) {
			System.out.println(e);
		}
		
		
	}
	
	public static void readexcel() {
		
		try {
			FileInputStream file = new FileInputStream("employees.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheet("records");
			
			//loop over rows in sheet
			
			Iterator<Row> rowiterator = sheet.rowIterator();
			while(rowiterator.hasNext()) {
				Row row = rowiterator.next();
				
				//loop over columns in each row
				
				Iterator<Cell> celliterator = row.cellIterator();
				while(celliterator.hasNext()) {
					Cell cell = celliterator.next();
					System.out.print(cell.getStringCellValue()+ "\t");
				} //end column loop
				System.out.print("\n");
			}//end ng rowloop
			
			file.close();
			System.out.println("tapos na");
		
			
		} catch (Exception e) {
			System.out.println(e);
		}
		
	}
	
	
	public static void createworkbook() {
		//write to xlsx
				//create instance workbook (cache instance sa ram)
				
				XSSFWorkbook workbook = new XSSFWorkbook();
				XSSFSheet sheet = workbook.createSheet("Employees"); //Employees name ng worksheet
				
				//data
				
				Map<String, Object[]> data = new TreeMap<String, Object[]>();
				data.put("1", new Object[] {"id", "name", "department"}); //object tatlo
				data.put("2", new Object[] {"1", "pepper", "hr"});
				data.put("3", new Object[] {"2", "chase", "accounting"});
				data.put("4", new Object[] {"3", "akli", "admin"});
				
				Set<String> keyset = data.keySet();
				
				int rownum=0;
				//loop
				for(String key:keyset) {
					
					Row row = sheet.createRow(rownum+=1);
					Object[] obj = data.get(key);
					int cellnum =0;
					
					//loop each column in each row
					
					for(Object o:obj) { 				//x3 lang tatakbo kasi id name department
						Cell cell = row.createCell(cellnum+=1);
						cell.setCellValue(o.toString());
					}//end of columnloop	
				}//end of row loop
				
				
				//write file in filesystem
						try {
							FileOutputStream out = new FileOutputStream(new File("employees.xlsx"));
							workbook.write(out);
							out.close();
							System.out.println("Write xlsx ok");
						} catch (Exception e) {
							System.out.println(e);
							}
	}
	
	public static void createworkbook(String workbookname, String worksheetname) {
		//write to xlsx
				//create instance workbook (cache instance sa ram)
				
				XSSFWorkbook workbook = new XSSFWorkbook();
				XSSFSheet sheet = workbook.createSheet(worksheetname); //Employees name ng worksheet
				
				//data
				
				Map<String, Object[]> data = new TreeMap<String, Object[]>();
				data.put("1", new Object[] {"id", "name", "department"}); //object tatlo
				data.put("2", new Object[] {"1", "pepper", "hr"});
				data.put("3", new Object[] {"2", "chase", "accounting"});
				data.put("4", new Object[] {"3", "akli", "admin"});
				
				Set<String> keyset = data.keySet();
				
				int rownum=0;
				//loop
				for(String key:keyset) {
					
					Row row = sheet.createRow(rownum+=1);
					Object[] obj = data.get(key);
					int cellnum =0;
					
					//loop each column in each row
					
					for(Object o:obj) { 				//x3 lang tatakbo kasi id name department
						Cell cell = row.createCell(cellnum+=1);
						cell.setCellValue(o.toString());
					}//end of columnloop	
				}//end of row loop
				
				
				//write file in filesystem
						try {
							FileOutputStream out = new FileOutputStream(new File(workbookname + ".xlsx"));
							workbook.write(out);
							out.close();
							System.out.println("Write xlsx ok");
						} catch (Exception e) {
							System.out.println(e);
							}
	}
}
