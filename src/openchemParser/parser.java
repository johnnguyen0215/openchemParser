package openchemParser;

import java.io.File; 
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfWriter;


public class parser {
	  public static void splitPDF(FileInputStream inputStream,
	          FileOutputStream outputStream, int fromPage, int toPage) {

		      Document document = new Document();
		
		      try {
		              PdfReader inputPDF = new PdfReader(inputStream);
		              int totalPages = inputPDF.getNumberOfPages();
		
		              // Make fromPage equals to toPage if it is greater
		              if (fromPage > toPage) {
		                      fromPage = toPage;
		              }
		              if (toPage > totalPages) {
		                      toPage = totalPages;
		              }
		
		              // Create a writer for the outputstream
		              PdfWriter writer = PdfWriter.getInstance(document, outputStream);
		              document.open();
		              // Holds the PDF data
		              PdfContentByte cb = writer.getDirectContent();
		              PdfImportedPage page;
		
		              while (fromPage <= toPage) {
		                      document.newPage();
		                      page = writer.getImportedPage(inputPDF, fromPage);
		                      cb.addTemplate(page, 0, 0);
		                      fromPage++;
		              }
		              outputStream.flush();
		              document.close();
		              outputStream.close();
		      } catch (Exception e) {
		              System.err.println(e.getMessage());
		      } finally {
		              if (document.isOpen())
		                      document.close();
		              try {
		                      if (outputStream != null)
		                              outputStream.close();
		              } catch (IOException ioe) {
		                      System.err.println(ioe.getMessage());
		              }
		      }
	}
	  
	public static void makePdfs(String pdfPages, String type, String rootDirName){
		if (pdfPages.contains(",")){
        	List<String> chunks = Arrays.asList(pdfPages.split(","));
        	for (int i = 0; i < chunks.size(); i++){
        		List<String> pdfRange = Arrays.asList(chunks.get(i).split("-"));
        		try{
        			splitPDF(new FileInputStream("openstax-chem.pdf"),
        					new FileOutputStream("D:/openchemPdfs/"+rootDirName
        							+"/"+type+"/"+(Integer.toString(i)+".pdf")), Integer.parseInt(pdfRange.get(0)), 
        							Integer.parseInt(pdfRange.get(1)));
        		}
        		catch (Exception e){
        			System.err.println(e.getMessage());
        		}
        	}
        }
		else{
			List<String> pdfRange = Arrays.asList(pdfPages.split("-"));
    		try{
    			splitPDF(new FileInputStream("openstax-chem.pdf"),
    					new FileOutputStream("D:/openchemPdfs/"+rootDirName
    							+"/"+type+"/"+"1.pdf"), Integer.parseInt(pdfRange.get(0)), 
    							Integer.parseInt(pdfRange.get(1)));
    		}
    		catch (Exception e){
    			System.err.println(e.getMessage());
    		}
		}
	}
	

	public static void main(String[] args) {
        try
        {
        	PrintWriter writer = new PrintWriter("laravelSeed.txt", "UTF-8");
        	
            FileInputStream file = new FileInputStream(new File("chemsubset.xlsx"));
 
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
                
                // Title
                Cell titleCell = row.getCell(0);
                String rootDirectoryName = titleCell.getStringCellValue().replaceAll("\\W", "");
                new File("D:\\openchemPdfs\\"+rootDirectoryName).mkdir();
                
                // Readings
                Cell readingsCell = row.getCell(3);
                System.out.println(readingsCell.getStringCellValue());
                if (readingsCell != null){
	                String readingsString = readingsCell.getStringCellValue();
	                new File("D:\\openchemPdfs\\"+rootDirectoryName+"\\Readings").mkdir();
	                makePdfs(readingsString, "Readings", rootDirectoryName);
                }
                
                // Problems
                Cell problemsCell = row.getCell(4);
                if (problemsCell != null){
	                String problemsString = problemsCell.getStringCellValue();
	        		new File("D:\\openchemPdfs\\"+rootDirectoryName+"\\Problems").mkdir();
	        		makePdfs(problemsString, "Problems", rootDirectoryName);
                }
                
        		// Solutions
                Cell solutionsCell = row.getCell(5);
                if (solutionsCell != null){
	                String solutionsString = solutionsCell.getStringCellValue();
	                new File("D:\\openchemPdfs\\"+rootDirectoryName+"\\Solutions").mkdir();
	                makePdfs(solutionsString, "Solutions", rootDirectoryName);    
                }
                
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



/*
for (int i = 0; i < 6; i++){
	Cell cell = row.getCell(i);
	if (row.getCell(i) != null){
		List<String> items;
    	switch (cell.getCellType())
    	{
	      case Cell.CELL_TYPE_NUMERIC:
	      	DataFormatter dataFormatter = new DataFormatter();
	      	String numVal = dataFormatter.formatCellValue(cell);
	      	System.out.print(numVal + " ");
	      	break;
		  case Cell.CELL_TYPE_STRING:
		  	String stringVal = cell.getStringCellValue();
		  	switch(i){
		  	case 0:
		  		rootDirectoryName = stringVal;
		  		new File("D:\\openchemPdfs\\"+rootDirectoryName).mkdir();
		  		break;
		  	case 3:
		  		new File("D:\\openchemPdfs\\"+rootDirectoryName+"\\Readings").mkdir();
		  	case 4:
		  		new File("D:\\openchemPdfs\\"+rootDirectoryName+"\\Problems").mkdir();
		  	case 5:
		  		new File("D:\\openchemPdfs\\"+rootDirectoryName+"\\Solutions").mkdir();
		  	}
		  	if (i == 0){
		  		stringVal = stringVal.replaceAll("\\W", "");
		  		System.out.print(stringVal + " ");
		  	}
		  	else if (i == 3){
		  		items = Arrays.asList(stringVal.split("-"));
		  		//System.out.print(Integer.parseInt(items.get(0))+1);
		  		for (int j = 0; j < items.size(); j++){
		  			
		  		}
		  	}
		  	else{
		  		System.out.print(stringVal + " ");
	      	}
	       break;
    	}
	}
}*/
/*
while (cellIterator.hasNext())
{
    Cell cell = cellIterator.next();
    //Check the cell type and format accordingly
  
    switch (cell.getCellType())
    {
        case Cell.CELL_TYPE_NUMERIC:
        	DataFormatter dataFormatter = new DataFormatter();
        	String numVal = dataFormatter.formatCellValue(cell);
        	System.out.print(numVal + " ");
        	break;
        case Cell.CELL_TYPE_STRING:
        	String stringVal = cell.getStringCellValue();
        	if (counter == 0){
        		stringVal = stringVal.replaceAll("\\W", "");
        		System.out.print(stringVal + " ");
        	}
        	else if (counter >= 3 && counter  <= 5){
        		List<String> items = Arrays.asList(stringVal.split("-"));
        		System.out.print(items + " ");
        	}
        	else{
        		System.out.print(stringVal + " ");
        	}
            
            
            break;
    }
    counter ++;
}
counter = 0;*/
