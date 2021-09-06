package poi_example.word_to_excel;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
 
public class sample {

	public static XWPFDocument docx; 
	public static XWPFWordExtractor we;
	
	public static void main(String[] args) {
		
		  try { docx = new XWPFDocument(new FileInputStream(
		  "C:\\Users\\jaswa\\Desktop\\munish_new\\MMAI\\capstone\\20210716 - CDA Team Meeting.docx"
		  ));
		  we = new XWPFWordExtractor(docx); 
		  System.out.println(we.getText());
		  }catch(Exception e) { e.printStackTrace(); }
		 

	}

}
