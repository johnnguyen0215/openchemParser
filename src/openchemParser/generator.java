package openchemParser;

import java.io.File;  
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class generator {
	
	public static String getVideoId(String url){
        String pattern = "(?<=watch\\?v=|/videos/|embed\\/)[^#\\&\\?]*";

        Pattern compiledPattern = Pattern.compile(pattern);
        Matcher matcher = compiledPattern.matcher(url);
        if(matcher.find()){
            return matcher.group();
        }
        return null;
	}
	
	public static String generateFileCode(String pdfPages, String type, String rootDirName){
		String code = "";
		if (pdfPages.contains(",")){
        	List<String> chunks = Arrays.asList(pdfPages.replaceAll("\\s", "").split(","));
        	for (int i = 0; i < chunks.size(); i++){
        		code += "$"+rootDirName + type + (i+1) + " = " + "Chemtext::create(array('chemtext_type' "
        				+ "=> 'pdf', 'chemtext_name' => 'OpenStax Chemistry',"
        				+ "'url' => " + "../uploads/Chem1A/"+type+"/"+((i+1)+".pdf));");
        		if (i+1 < chunks.size()){
        			code += "\n";
        		}
        	}
		}
		else{
			code += "$"+rootDirName + type + " = " + "Chemtext::create(array('chemtext_type' "
    				+ "=> 'pdf', 'chemtext_name' => 'OpenStax Chemistry',"
    				+ "'url' => " + "../uploads/Chem1A/"+type+"/"+"1.pdf));";
		}
		
		return code;
		
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
            
            DataFormatter dataFormatter = new DataFormatter();
            
            String chem1ADescription = "Description: Chem 1A is the first quarter of General "
            		+ "Chemistry and covers the following topics: atomic structure; "
            		+ "general properties of the elements; covalent, ionic, and metallic bonding; "
            		+ "intermolecular forces; mass relationships. General Chemistry (Chem 1A) is part of "
            		+ "OpenChemThis video is part of a 23-lecture undergraduate-level course titled "
            		+ "\"General Chemistry\" taught at UC Irvine by Amanda Brindley, Ph.D."; 

            while (rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                
                Cell titleCell = row.getCell(0);
                Cell videoInCell = row.getCell(1);
                Cell videoOutCell = row.getCell(2);
                Cell readingsCell = row.getCell(3, Row.RETURN_BLANK_AS_NULL);
        		Cell problemsCell = row.getCell(4, Row.RETURN_BLANK_AS_NULL);
				Cell solutionsCell = row.getCell(5, Row.RETURN_BLANK_AS_NULL);
                Cell videoUrlCell = row.getCell(6);
                
                String title = titleCell.getStringCellValue();
                String videoUrl = videoUrlCell.getStringCellValue();
                String videoId = getVideoId(videoUrl);
                String videoIn = dataFormatter.formatCellValue(videoInCell);
                String videoOut = dataFormatter.formatCellValue(videoOutCell);
                
                String youtubeUrl = "https://www.youtube.com/watch?start="+
                videoIn+"&end="+videoOut+"&v="+videoId;
                
                if (readingsCell != null){
	                String readingsString = dataFormatter.formatCellValue(readingsCell);
	                writer.println(generateFileCode(readingsString, "Readings",
	                		title.replaceAll("\\W", "")));
                }
                
                if (problemsCell != null){
                	String problemsString = dataFormatter.formatCellValue(problemsCell);
	                writer.println(generateFileCode(problemsString, "Problems",
	                		title.replaceAll("\\W", "")));
                }
                
                if (solutionsCell != null){
                	String solutionsString = dataFormatter.formatCellValue(solutionsCell);
	                writer.println(generateFileCode(solutionsString, "Solutions",
	                		title.replaceAll("\\W", "")));
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
