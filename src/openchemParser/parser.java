package openchemParser;

import java.io.File;
import java.io.FileInputStream;
import java.io.PrintWriter;
import java.util.Iterator;

import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class parser {

	public static void main(String[] args) {
        try
        {
        	PrintWriter writer = new PrintWriter("laravelSeed.txt", "UTF-8");
        	
            FileInputStream file = new FileInputStream(new File("Chem 1A.xlsx"));
 
            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);
 
            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
 
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                Iterator<Cell> cellIterator = row.cellIterator();
                
                String topicName;
                String videoId;
                String videoUrl;
                String videoDescription; 
                
                while (cellIterator.hasNext())
                {
                    Cell cell = cellIterator.next();
                    //Check the cell type and format accordingly
                  
                    switch (cell.getCellType())
                    {
                        case Cell.CELL_TYPE_NUMERIC:
                        	DataFormatter dataFormatter = new DataFormatter();
                        	String val = dataFormatter.formatCellValue(cell);
                        	
                        	break;
                        case Cell.CELL_TYPE_STRING:
                            System.out.print(cell.getStringCellValue() + " ");
                            writer.print(cell.getStringCellValue()+ " ");
                            break;
                    }
                }
                System.out.println("");
                writer.println("");
            }
            file.close();
            writer.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }		
		
		
	}

}
