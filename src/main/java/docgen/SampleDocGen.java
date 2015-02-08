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
import org.apache.poi.ss.usermodel.Workbook;

import net.sf.jett.transform.ExcelTransformer;

public class SampleDocGen {

	public static void main(String[] args) throws InvalidFormatException, IOException {
		// TODO 自動生成されたメソッド・スタブ
		SampleDocGen.makeShift();
	}

	public static File makeShift() throws InvalidFormatException, IOException {
		System.out.println(System.getProperty("user.dir"));

	 HashMap beans = new HashMap();
	 ArrayList<HashMap> records =  new ArrayList<HashMap>();
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
		//Map<String, Object> beans;
		//Workbook workbook = transformer.transform(fileIn, beans);
		//Workbook workbook = transformer.transform(fileIn, beans);
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
	   	vbeans.put("models",list);
//	   	 vbeans.put("ht002", "222222");        
        beansList.add(vbeans);
		Workbook workbook = transformer.transform(fileIn, templateSheetNames, sheetNames, (List<Map<String, Object>>) beansList);
		//		transformer.ttransform(fileIn, beans);
		 workbook.write(fileOut);
		fileOut.close();
		 return new File(outPath);
	}
}
