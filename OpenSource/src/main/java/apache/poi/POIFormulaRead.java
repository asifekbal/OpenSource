package apache.poi;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POIFormulaRead {

	public static void main(String[] args) {
		try {
			FileInputStream file = new FileInputStream("ApachePOIFormulaExample.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
			Iterator<Row> rowIterator = sheet.iterator();
			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				Iterator<Cell> cellIterator = row.iterator();
				while (cellIterator.hasNext()) {
					Cell cell = cellIterator.next();
					switch (evaluator.evaluateInCell(cell).getCellType()) {
					case NUMERIC:
						System.out.print(cell.getNumericCellValue() + "\t");
						break;
					case STRING:
						System.out.print(cell.getStringCellValue() + "\t");
						break;
					case FORMULA:
						break;
					default:
						break;
					}
				}
				System.out.println("");
			}
			file.close();
			workbook.close();

		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
