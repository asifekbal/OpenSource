package apache.poi;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POIFormulaWrite {

	public static void main(String[] args) {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Calculate Simple Interest");

		Row header = sheet.createRow(0);
		header.createCell(0).setCellValue("Principal");
		header.createCell(1).setCellValue("RoI");
		header.createCell(2).setCellValue("T");
		header.createCell(3).setCellValue("Interest (P r t)");

		Row dataRow = sheet.createRow(1);
		dataRow.createCell(0).setCellValue(14500d);
		dataRow.createCell(1).setCellValue(9.25);
		dataRow.createCell(2).setCellValue(3d);
		dataRow.createCell(3).setCellFormula("A2*B2*C2");

		try (FileOutputStream out = new FileOutputStream(new File("ApachePOIFormulaExample.xlsx"));) {
			workbook.write(out);
			workbook.close();
			System.out.println("File successfully written on disk");
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}
}