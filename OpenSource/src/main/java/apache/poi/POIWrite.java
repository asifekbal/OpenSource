package apache.poi;

import java.io.File;
import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POIWrite {

	public static void main(String[] args) {
		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Employee Data");
		Map<String, Object[]> data = new HashMap<>();
		data.put("1", new Object[] { "ID", "NAME", "SALARY" });
		data.put("2", new Object[] { 1, "Asif", 5000 });
		data.put("3", new Object[] { 2, "Rahul", 8000 });
		data.put("4", new Object[] { 3, "Keshav", 4000 });
		data.put("5", new Object[] { 4, "Deepak", 7500 });

		Set<String> keyset = data.keySet();
		int rownum = 0;
		for (String key : keyset) {
			Row row = sheet.createRow(rownum++);
			Object[] objArr = data.get(key);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof String) {
					cell.setCellValue((String) obj);
				} else if (obj instanceof Integer) {
					cell.setCellValue((Integer) obj);
				}
			}
		}
		try (FileOutputStream out = new FileOutputStream(new File("ApachePOIExample.xlsx"));) {
			workbook.write(out);
			System.out.println("File successfully written on disk");
			workbook.close();
		} catch (Exception ex) {
			ex.printStackTrace();
		}
	}

}
