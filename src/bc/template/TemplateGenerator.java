package bc.template;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * class TemplateGenerator:
 *  - To maintain the existence of comparison template which is used as an input
 *	for comparing process.
 *
 * */
public class TemplateGenerator {
	final private static String TEMPLATE_LOCATION = System.getProperty("user.dir") + "/input/";
	final private static String TEMPLATE_FILENAME = "CompareTemplate.xls";
	
	private static boolean checkTemplateExist() {
		boolean exist = false;
		File template = new File(TEMPLATE_LOCATION + TEMPLATE_FILENAME);
		if(template.exists() && !template.isDirectory()) {
			exist = true;
		}
		return exist;
	}
	
	public static void createTemplate() {
		if(checkTemplateExist() == false) {
			// create input directory
			File f = new File(TEMPLATE_LOCATION);
			f.mkdir();
			
			// create blank Workbook (excel file)
			Workbook wb = new XSSFWorkbook();
			try(OutputStream out = new FileOutputStream(TEMPLATE_LOCATION + TEMPLATE_FILENAME)) {
				wb.write(out);
				out.close();
				wb.close();
			} catch (Exception e) {
				System.out.println("Exception: " + e.getMessage());
			}
		}
	}
}
