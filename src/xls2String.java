import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import jxl.read.biff.BiffException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;


public class xls2String {
	
	private int rows;
	private int cols;
	private String stringArray[][]; 
	

	public xls2String(){
	}
	
	public xls2String(String filePath){
		
		int i = 0; int j = 0;
		
		setCols(filePath);
		setRows(filePath);
		String data[][] = new String[rows][cols]; 
		
        try {

            FileInputStream file = new FileInputStream(new File(filePath));
            
            // Get the workbook instance for XLS file
            HSSFWorkbook workbook = new HSSFWorkbook(file);

            // Get first sheet from the workbook
            HSSFSheet sheet = workbook.getSheetAt(0);
            
            // Iterate through each rows from first sheet
            Iterator<Row> rowIterator = sheet.iterator();
            
            while (rowIterator.hasNext() && i < rows) {
                Row row = rowIterator.next();
                j = 0;

                // For each row, iterate through each columns
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()  && j < cols) {

                    Cell cell = cellIterator.next();

                    switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_BOOLEAN:
                        data[i][j] = String.valueOf(cell.getBooleanCellValue());
                        j++;
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                    	 data[i][j] = String.valueOf(cell.getNumericCellValue());
                    	 j++;
                        break;
                    case Cell.CELL_TYPE_STRING:
                    	data[i][j] = cell.getStringCellValue();
                    	j++;
                        break;
                    }
                }
                //Goto next row
                i++;

            }
            workbook.close();
            file.close();
            
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
        
        stringArray = data;
	}
	
	

	public int getCols() {
		return cols;
	}
	
	

	private void setCols(String filePath) {
		try {
			jxl.Workbook workbook = jxl.Workbook.getWorkbook(new File (filePath));
			jxl.Sheet sheet = workbook.getSheet(0);
			cols = sheet.getColumns();
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	


	public int getRows() {
		return rows;
	}

	
	
	private void setRows(String filePath) {
		try {
			jxl.Workbook workbook = jxl.Workbook.getWorkbook(new File (filePath));
			jxl.Sheet sheet = workbook.getSheet(0);
			rows = sheet.getRows();
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public String[][] getStringArray(){
		return stringArray;
	}
	
	public static void main(String[] args) {
		
		String file = "test_file/test.xls";
		xls2String data = new xls2String(file);
		
		String arr[][] = data.getStringArray();
		
		System.out.println(data.getRows());
		System.out.println(data.getCols());

		for (int i = 0; i < arr.length; i++){
			for(int j = 0; j < arr[i].length; j++){
				System.out.print(arr[i][j] + " \t");
			}
			System.out.println();
		}
	}
	
	
}
