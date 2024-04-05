package testapp1;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoiDemo {

	public static void main(String[] args) {
//		createworkbook("employee", "records");
//		readexcel("employees", "records");
		appendrow("employees","records", "4", "lakan", "general");
	}
		//append row
		public static void appendrow(String wb, String ws, String id, String name, String department) {
		
			try {
				File file1 = new File(wb + ".xlsx");
				//check if existing
				FileInputStream file = new FileInputStream(file1);
		
				if(file1.exists()) {
					
					XSSFWorkbook workbook = new XSSFWorkbook(wb + ".xlsx");
					XSSFSheet sheet = workbook.getSheet(ws);
					int rowlastnum = sheet.getLastRowNum();
					Row newrow = sheet.createRow(rowlastnum + 1);
					Cell cell1 = newrow.createCell(0);
					cell1.setCellValue(id);
					Cell cell2 = newrow.createCell(1);
					cell2.setCellValue(name);
					Cell cell3 = newrow.createCell(2);
					cell3.setCellValue(department);
					
					//write to file
					FileOutputStream out = new FileOutputStream(wb + ".xlsx");
					workbook.write(out);
					System.out.println("new row added");
					out.close();
			
				}else {
					System.out.println("cannot append, file not exist");
				}
				

			} catch (Exception e) {
				System.out.println(e);
			}
			
	}
	
		public static void readexcel(String workbookname, String worksheetname) {
			
			try {
				File file = new File(workbookname + ".xlsx");
				XSSFWorkbook workbook = new XSSFWorkbook(file);
				XSSFSheet sheet = workbook.getSheet(worksheetname);
				
				//loop over rows in sheet
				Iterator<Row> rowiterator = sheet.rowIterator();
				while(rowiterator.hasNext()) {
					Row row = rowiterator.next();
					
					//loop over columns in each row
					Iterator<Cell> celliterator = row.cellIterator();
					while(celliterator.hasNext()) {
						Cell cell =celliterator.next();
						System.out.print(cell.getStringCellValue() + "\t");
					}
					System.out.print("\n");
				}
				
			} catch (InvalidFormatException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			
			} catch (IOException e) {
			// TODO Auto-gene();
		
		}
		}
			
		public static void createworkbook(String workbookname, String worksheetname) {
		
		//write to xlsx
		//create instance workbook
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("records");
		
		//data
		Map<String, Object[]> data = new TreeMap<String, Object[]>();
		data.put("1", new Object[] {"id", "name", "department"});
		data.put("2", new Object[] {"1", "akio", "HR"});
		data.put("3", new Object[] {"2", "maomao", "Medical"});
		data.put("4", new Object[] {"3", "jinshi", "Management"});
		
		Set<String> keyset = data.keySet();
		
		int rownum =0;
		//loop each column in each row
		for(String key:keyset) {
			
			Row row = sheet.createRow(rownum+=1);
			Object[] obj = data.get(key);
			int cellnum =0;
			for(Object o:obj) {
				Cell cell = row.createCell(cellnum++);
				cell.setCellValue(o.toString());	
			}
			//write file in the filesystem
			
		}
		
	
		try {
			//File file = new File("file.xlsx")
			FileOutputStream out = new FileOutputStream(new File("employees.xlsx"));
			workbook.write(out);
			out.close();
			System.out.println("write xlsx ok");	
		} catch (Exception e) {
			System.out.println(e);
		}
		
		
		
	}
}
