package openchemParser;

import java.io.File;  
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
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
	
	  public static int rowCounter = 1;
	
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
		String inFile = "openstax-chem.pdf";
		int rangeFrom;
		int rangeTo;
		if (pdfPages.contains(",")){
        	List<String> chunks = Arrays.asList(pdfPages.replaceAll("\\s", "").split(","));
        	for (int i = 0; i < chunks.size(); i++){
        		List<String> pdfRange = Arrays.asList(chunks.get(i).split("-"));
        		try{
        			String outFile = "C:/openchemPdfs/"+rootDirName
							+"/"+type+"/"+(Integer.toString(i+1)+".pdf");
        			rangeFrom = Integer.parseInt(pdfRange.get(0));
        			rangeTo = Integer.parseInt(pdfRange.get(1));
        			splitPDF(new FileInputStream(inFile),
        					new FileOutputStream(outFile), rangeFrom, rangeTo);
        		}
        		catch (Exception e){
        			System.out.println("First Exception, Row Number: " + rowCounter);
        			e.printStackTrace();
        		}
        	}
        }
		else{
			List<String> pdfRange = Arrays.asList(pdfPages.replaceAll("\\s", "").split("-"));
    		try{
    			
    			String outFile = "C:/openchemPdfs/" + rootDirName
    					+ "/" + type + "/" + "1.pdf";
    			if (pdfRange.size() > 1){
	    			rangeFrom = Integer.parseInt(pdfRange.get(0));
	    			rangeTo = Integer.parseInt(pdfRange.get(1));
    			}
    			else{
    				rangeFrom = rangeTo = Integer.parseInt(pdfRange.get(0));
    			}
    			//System.out.print(outFile + " " + pdfRange.get(0) + " " + pdfRange.get(1));
    			splitPDF(new FileInputStream(inFile),
    					new FileOutputStream(outFile), rangeFrom, 
    							rangeTo);
    		}
    		catch (Exception e){
    			System.out.println("Second Exception, Row Number: " + rowCounter);
    			e.printStackTrace();
    		}
		}
	}
	

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
            
            DataFormatter dataFormatter = new DataFormatter();
            while (rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                
                // Title
                Cell titleCell = row.getCell(0);
                Cell readingsCell = row.getCell(3, Row.RETURN_BLANK_AS_NULL);
                Cell problemsCell = row.getCell(4, Row.RETURN_BLANK_AS_NULL);
                Cell solutionsCell = row.getCell(5, Row.RETURN_BLANK_AS_NULL);
                String rootDirectoryName = titleCell.getStringCellValue().replaceAll("\\W", "");
                
                if (readingsCell != null || problemsCell != null || solutionsCell != null){
                	new File("C:\\openchemPdfs\\"+rootDirectoryName).mkdir();
                }
                
                // Readings
                if (readingsCell != null){
	                String readingsString = dataFormatter.formatCellValue(readingsCell);
	                new File("C:\\openchemPdfs\\"+rootDirectoryName+"\\Readings").mkdir();
	                makePdfs(readingsString, "Readings", rootDirectoryName);
                }
                
                // Problems
                if (problemsCell != null){
	                String problemsString = dataFormatter.formatCellValue(problemsCell);
	        		new File("C:\\openchemPdfs\\"+rootDirectoryName+"\\Problems").mkdir();
	        		makePdfs(problemsString, "Problems", rootDirectoryName);
                }
                
        		// Solutions
                if (solutionsCell != null){
	                String solutionsString = dataFormatter.formatCellValue(solutionsCell);
	                new File("C:\\openchemPdfs\\"+rootDirectoryName+"\\Solutions").mkdir();
	                makePdfs(solutionsString, "Solutions", rootDirectoryName);    
                }
                rowCounter ++;
 
            }
            file.close();
            writer.close();
        }
        catch (Exception e)
        {
        	System.out.println("Third Exception, Row Number: " + rowCounter);
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
		  		new File("C:\\openchemPdfs\\"+rootDirectoryName).mkdir();
		  		break;
		  	case 3:
		  		new File("C:\\openchemPdfs\\"+rootDirectoryName+"\\Readings").mkdir();
		  	case 4:
		  		new File("C:\\openchemPdfs\\"+rootDirectoryName+"\\Problems").mkdir();
		  	case 5:
		  		new File("C:\\openchemPdfs\\"+rootDirectoryName+"\\Solutions").mkdir();
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
