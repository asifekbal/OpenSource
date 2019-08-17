package apache.poi;

import java.io.File;
import java.io.FileOutputStream;

import org.apache.poi.hssf.record.CFRuleBase.ComparisonOperator;
import org.apache.poi.ss.usermodel.ConditionalFormattingRule;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PatternFormatting;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POIFormattingWrite {

	public static void main(String[] args) {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Cell Formatting");
		try (FileOutputStream out = new FileOutputStream(new File("ApachePOIFormattingCell.xlsx"));) {
			basedOnValue(sheet);
			workbook.write(out);
			System.out.println("File successfully written on disk");
			workbook.close();
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

	private static void basedOnValue(XSSFSheet sheet) {

		sheet.createRow(0).createCell(0).setCellValue(84);
		sheet.createRow(1).createCell(0).setCellValue(74);
		sheet.createRow(2).createCell(0).setCellValue(50);
		sheet.createRow(3).createCell(0).setCellValue(51);
		sheet.createRow(4).createCell(0).setCellValue(49);
		sheet.createRow(5).createCell(0).setCellValue(41);

		SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();

		ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule(ComparisonOperator.GE, "70");
		PatternFormatting fill1 = rule1.createPatternFormatting();
		fill1.setFillBackgroundColor(IndexedColors.BLUE.index);
		fill1.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

		ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(ComparisonOperator.LT, "50");
		PatternFormatting fill2 = rule2.createPatternFormatting();
		fill2.setFillBackgroundColor(IndexedColors.GREEN.index);
		fill2.setFillPattern(PatternFormatting.SOLID_FOREGROUND);

		CellRangeAddress[] regions = { CellRangeAddress.valueOf("A1:A6") };

		sheetCF.addConditionalFormatting(regions, rule1, rule2);

	}

}
