package docgen;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import net.sf.jett.transform.ExcelTransformer;

public class SampleDocGen {

	public static void main(String[] args) throws InvalidFormatException,
			IOException {
		// TODO 自動生成されたメソッド・スタブ
		//SampleDocGen.makeShift();
		SampleDocGen.setColor();
	}

	public static void setColor() {

		try {
			System.out.println(System.getProperty("user.dir"));

			FileInputStream fi = new FileInputStream("1.xlsx");
			Workbook book = new XSSFWorkbook(fi);
			// Workbook book =new HSSFWorkbook(fi);
			fi.close();
			// for(Sheet sheet:book){ // XSSFWorkbookの場合
			CellStyle style = book.createCellStyle();

			for (int s = 0; s < book.getNumberOfSheets(); ++s) { // 全シートをなめる(※)
				Sheet sheet = book.getSheetAt(s);

				
				// sheet.setForceFormulaRecalculation(true); // 数式解決(※２)
				System.out.println("--- " + sheet.getSheetName() + " ---");
				for (Row row : sheet) { // 全行をなめる
					Cell cell2 = row.createCell(1);
					for (Cell cell : row) { // 全セルをなめる
//						style.setFillBackgroundColor(IndexedColors.PINK.getIndex());
//						style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
//						style.setFillPattern(CellStyle.SOLID_FOREGROUND); // 塗りつぶし
					    style.setFillPattern(CellStyle.BIG_SPOTS);
					    style.setFillForegroundColor(IndexedColors.RED.getIndex());
					    style.setFillBackgroundColor(IndexedColors.WHITE.getIndex());
						cell.setCellStyle(style);
					}
				    style.setFillPattern(CellStyle.BIG_SPOTS);
				    style.setFillForegroundColor(IndexedColors.RED.getIndex());
				    style.setFillBackgroundColor(IndexedColors.WHITE.getIndex());					
					cell2.setCellStyle(style);
					System.out.println("");
				}
			}
			String outPath = "2.xls";
			FileOutputStream fileOut = null;
			try {
				fileOut = new FileOutputStream(outPath);
			} catch (IOException e) {
				System.err.println("IOException opening " + outPath + ": "
						+ e.getMessage());
			}
			book.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace(System.err);
			System.exit(1);
		}
		System.exit(0);
	}

	public static File makeShift() throws InvalidFormatException, IOException {

		HashMap beans = new HashMap();
		ArrayList<HashMap> records = new ArrayList<HashMap>();
		HashMap vbeans = new HashMap();
		vbeans.put("ht001", "111111");
		vbeans.put("ht002", "111111");
		vbeans.put("ht003", "111111");
		vbeans.put("ht004", "111111");
		records.add(vbeans);
		records.add(vbeans);
		records.add(vbeans);

		beans.put("test_val", "aa");
		beans.put("models", records);

		String inPath = "template.xls";
		String outPath = "shift.xls";
		FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream(outPath);
		} catch (IOException e) {
			System.err.println("IOException opening " + outPath + ": "
					+ e.getMessage());
		}
		InputStream fileIn = null;
		fileIn = new BufferedInputStream(new FileInputStream(inPath));
		ExcelTransformer transformer = new ExcelTransformer();
		// Map<String, Object> beans;
		// Workbook workbook = transformer.transform(fileIn, beans);
		// Workbook workbook = transformer.transform(fileIn, beans);
		List<String> templateSheetNames = new ArrayList<String>();
		templateSheetNames.add("目次");
		List<String> sheetNames = new ArrayList<String>();
		sheetNames.add("目次");
		List<Map<String, Object>> beansList = new ArrayList<Map<String, Object>>();
		HashMap vbean2 = new HashMap();
		vbean2.put("ht001", "1");
		vbean2.put("ht002", "2");
		vbean2.put("ht003", "3");
		vbean2.put("ht004", "4");
		vbean2.put("ht005", "5");
		vbean2.put("ht006", "6");

		ArrayList list = new ArrayList();
		SummaryResult sr = new SummaryResult();
		sr.setHT001("134");
		list.add(vbean2);
		list.add(vbean2);
		list.add(vbean2);
		vbeans.put("models", list);
		// vbeans.put("ht002", "222222");
		beansList.add(vbeans);
		Workbook workbook = transformer.transform(fileIn, templateSheetNames,
				sheetNames, (List<Map<String, Object>>) beansList);
		// transformer.ttransform(fileIn, beans);
		workbook.write(fileOut);
		fileOut.close();
		return new File(outPath);
	}

	public static String getStr(Cell cell) { // データ型毎の読み取り
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_BOOLEAN:
			return Boolean.toString(cell.getBooleanCellValue());
		case Cell.CELL_TYPE_FORMULA:
			return cell.getCellFormula();
			// return cell.getStringCellValue();(※）
		case Cell.CELL_TYPE_NUMERIC:
			return Double.toString(cell.getNumericCellValue());
		case Cell.CELL_TYPE_STRING:
			return cell.getStringCellValue();
		}
		return "";// CELL_TYPE_BLANK,CELL_TYPE_ERROR
	}
}
