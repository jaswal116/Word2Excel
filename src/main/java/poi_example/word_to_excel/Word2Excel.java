package poi_example.word_to_excel;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

public class Word2Excel {


	static String[] HEADERs = { "Text", "Speaker", "Timestamp", "Psychological Safety", "Challenge", "Support", "Candour" };
	static String SHEET = " Data ";

	public static void main(String[] args) throws IOException {

		try (XWPFDocument docx = new XWPFDocument(new FileInputStream(
				"C:\\Users\\jaswa\\Desktop\\munish_new\\MMAI\\capstone\\20210716 - CDA Team Meeting.docx"));
				XWPFWordExtractor we = new XWPFWordExtractor(docx);
				XSSFWorkbook workbook = new XSSFWorkbook();
				OutputStream out = new FileOutputStream(
						"C:\\Users\\jaswa\\Desktop\\munish_new\\MMAI\\capstone\\Speak1.xlsx");
				ByteArrayOutputStream out2 = new ByteArrayOutputStream();) {
			// create a new Sheet using Workbook.createSheet()
			Sheet sheet = workbook.createSheet(SHEET);

			// Create a Font for styling header cells
			Font headerFont = workbook.createFont();
			headerFont.setBold(true);
			headerFont.setFontHeightInPoints((short) 14);
			headerFont.setColor(IndexedColors.BLACK.getIndex());

			// Create a CellStyle with the font
			CellStyle headerCellStyle = workbook.createCellStyle();
			headerCellStyle.setFont(headerFont);

			// Create a Row
			Row headerRow = sheet.createRow(0);

			// Create Header cells
			for (int i = 0; i < HEADERs.length; i++) {
				Cell cell = headerRow.createCell(i);
				cell.setCellValue(HEADERs[i]);
				cell.setCellStyle(headerCellStyle);
			}

			// Resize all columns to fit the content size
			for (int i = 0; i < HEADERs.length; i++) {
				sheet.autoSizeColumn(i);
			}
			int ro  = 1;

			List<XWPFParagraph> paragraphList = docx.getParagraphs();
			for (XWPFParagraph p : paragraphList) {

				Row row = sheet.createRow(ro++);

				int t = 0;

				Cell cell = row.getCell(t);
				if (cell == null) {
					cell = row.createCell(t);
				}
				//cell.setCellType(CellType.STRING);
				cell.setCellValue(p.getText());
				System.out.println(p.getText());
			}
			workbook.write(out);
			out.close();
			System.out.println("final completed");
			// Closing the workbook
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

}
