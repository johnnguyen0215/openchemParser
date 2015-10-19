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
	
	public static ArrayList<String> generateFileCode(String pdfPages, String dirName, String rootDirName, String course, PrintWriter writer){
		String type = "";
		String typeName = "";
		switch(dirName){
			case "Readings":
				type = "Chemtext";
				typeName = "chemtext_name";
				break;
			case "Problems":
				type = "Problem";
				typeName = "problem_name";
				break;
			case "Solutions":
				type = "Solution";
				typeName = "solution_name";
				break;
		}
		ArrayList<String> variables = new ArrayList<String>();
		if (pdfPages.contains(",")){
        	List<String> chunks = Arrays.asList(pdfPages.replaceAll("\\s", "").split(","));
        	for (int i = 0; i < chunks.size(); i++){
        		String code = "$"+rootDirName + dirName + (i+1) + " = " + type + "::create(array('" + typeName + "' => 'OpenStax Chemistry', "
        				+ "'url' => " + "\"../uploads/"+course+"/" + rootDirName + "/" + dirName+"/"+((i+1)+".pdf\"));");
        		writer.println(code);
        		variables.add("$" + rootDirName + dirName + (i+1));
        	}
		}
		else{
			String code = "$"+rootDirName + dirName + " = " + type + "::create(array('" + typeName + "' => 'OpenStax Chemistry', "
    				+ "'url' => " + "\"../uploads/"+course+"/"+ rootDirName + "/" + dirName+"/"+"1.pdf\"));";
			writer.println(code);
			variables.add("$" + rootDirName + dirName);
		}
		
		return variables;
		
	}
	
	public static String generateTopicCode(String course,String title, String videoUrl, String videoId, String videoDescription, PrintWriter writer){
		String code = "$" + course.replaceAll("\\W", "") + title.replaceAll("\\W", "") + " = " + "Topic::create(array('topic_name' => \"" + title + "\", 'video_url' => '"
		+ videoUrl + "', 'video_id' => \"" + videoId + "\", 'video_description' => \"" + videoDescription + "\"));";
		writer.println(code);
		
		return "$" + title.replaceAll("\\W", "");
	}
	
	public static void generateAttachmentsCode(String titleVar, ArrayList<String> readingsVars, ArrayList<String> problemsVars, ArrayList<String> solutionsVars,
			PrintWriter writer){
		String code;
		for (String var : readingsVars){
			code = titleVar + "->chemtexts()->attach(" + var + "->id);";
			writer.println(code);
		}
		for (String var : problemsVars){
			code = titleVar + "->problems()->attach(" + var + "->id);";
			writer.println(code);
		}
		for (String var : solutionsVars){
			code = titleVar + "->solutions()->attach(" + var + "->id);";
			writer.println(code);
		}
	}
	
	public static void initializeCodeGen(String course){
        try
        {
        	PrintWriter writer = new PrintWriter(course + " Seed.txt", "UTF-8");
        	
            FileInputStream file = new FileInputStream(new File(course+".xlsx"));
 
            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);
 
            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
 
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = sheet.iterator();
            
            DataFormatter dataFormatter = new DataFormatter();
            
            String videoDescription = "";
            
            switch (course) {
	            case "Chem 1A":
	            	videoDescription = "Chem 1A is the first quarter of General "
	                		+ "Chemistry and covers the following topics: atomic structure; "
	                		+ "general properties of the elements; covalent, ionic, and metallic bonding; "
	                		+ "intermolecular forces; mass relationships. General Chemistry (Chem 1A) is part of "
	                		+ "OpenChem. This video is part of a 23-lecture undergraduate-level course titled "
	                		+ "'General Chemistry' taught at UC Irvine by Amanda Brindley, Ph.D."; 
	            	break;
	            case "Chem 1B":
	            	videoDescription = "UCI Chem 1B is the second quarter of General "
	                		+ "Chemistry and covers the following topics: properties of gases, liquids, solids; "
	                		+ "changes of state; properties of solutions; stoichiometry; thermochemistry; and thermodynamics."
	                		+ "General Chemistry (Chem 1B) is part of OpenChem. This video is part of a 17-lecture "
	                		+ "undergraduate-level course titled 'General Chemistry' taught at UC Irvine by "
	                		+ "Donald R. Blake, Ph.D.";
	            /*
	            case "Chem 1C":
	            	videoDescription = "UCI Chem 1C is the third and final quarter of General Chemistry "
	                		+ "series and covers the following topics: equilibria, aqueous acid-base equilibria, solubility equilibria, "
	                		+ "oxidation reduction reactions, electrochemistry; kinetics; special topics. "
	                		+ "General Chemistry (Chem 1C) is part of OpenChem. This video is part of a 26-lecture undergraduate-level "
	                		+ "course titled 'General Chemistry' taught at UC Irvine by Ramesh D. Arasasingham, Ph.D.";*/
            }
            
            while (rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                //For each row, iterate through all the columns
                
                Cell titleCell = row.getCell(0, Row.RETURN_BLANK_AS_NULL);
                Cell videoInCell = row.getCell(1, Row.RETURN_BLANK_AS_NULL);
                Cell videoOutCell = row.getCell(2, Row.RETURN_BLANK_AS_NULL);
                Cell readingsCell = row.getCell(3, Row.RETURN_BLANK_AS_NULL);
        		Cell problemsCell = row.getCell(4, Row.RETURN_BLANK_AS_NULL);
				Cell solutionsCell = row.getCell(5, Row.RETURN_BLANK_AS_NULL);
                Cell videoUrlCell = row.getCell(6);
                
                String title = titleCell.getStringCellValue();
                String videoUrl = videoUrlCell.getStringCellValue();
                String videoId = getVideoId(videoUrl);
                String videoIn = dataFormatter.formatCellValue(videoInCell);
                
                String youtubeUrl = "";
                
                if (videoOutCell == null){
                	youtubeUrl = "https://www.youtube.com/watch?start="+
                            videoIn+"&v="+videoId;
                }
                else {
                	String videoOut = dataFormatter.formatCellValue(videoOutCell);
                    youtubeUrl = "https://www.youtube.com/watch?start="+
                            videoIn+"&end="+videoOut+"&v="+videoId;
                }
                
                
                ArrayList<String> readingsVariables = new ArrayList<String>();
                ArrayList<String> problemsVariables = new ArrayList<String>();
                ArrayList<String> solutionsVariables = new ArrayList<String>();
                
                if (readingsCell != null){
	                String readingsString = dataFormatter.formatCellValue(readingsCell);
	                readingsVariables = generateFileCode(readingsString, "Readings",
	                		title.replaceAll("\\W", ""), course, writer);
                }
                
                if (problemsCell != null){
                	String problemsString = dataFormatter.formatCellValue(problemsCell);
                	problemsVariables = generateFileCode(problemsString, "Problems",
	                		title.replaceAll("\\W", ""), course, writer);
                }
                
                if (solutionsCell != null){
                	String solutionsString = dataFormatter.formatCellValue(solutionsCell);
	                solutionsVariables = generateFileCode(solutionsString, "Solutions",
	                		title.replaceAll("\\W", ""), course, writer);
                }
                
                String topicVariable = generateTopicCode(course, title, youtubeUrl, videoId, videoDescription, writer);
                
                generateAttachmentsCode(topicVariable, readingsVariables, problemsVariables, solutionsVariables, writer);
                
                
            }
            file.close();
            writer.close();
            workbook.close();
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }		
            
	}
	
	
	public static void main(String[] args) {
		initializeCodeGen("Chem 1A");
		//initializeCodeGen("Chem 1B");
		//initializeCodeGen("Chem 1C");
	}

}
