package openchemParser;


import java.io.FileOutputStream;
import java.io.File; 
import java.io.FileInputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.util.Iterator;




import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;





import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfCopy;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfWriter;
import com.itextpdf.text.pdf.codec.Base64.InputStream;
import com.itextpdf.text.pdf.codec.Base64.OutputStream;

public class SplitPDFFile {

    /**
     * @param args
     */
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
	
    public static void main(String[] args) {
    	try{
	        splitPDF(new FileInputStream("math.pdf"),
	                new FileOutputStream("sample.pdf"), 3, 5);
    	}
    	catch (Exception e){
    		System.err.println(e.getMessage());
    	}
    }
}