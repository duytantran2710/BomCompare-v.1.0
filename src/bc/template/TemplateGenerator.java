package bc.template;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
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
	
	/*
	private static boolean checkTemplateExist() {
		boolean exist = false;
		File template = new File(TEMPLATE_LOCATION + TEMPLATE_FILENAME);
		if(template.exists() && !template.isDirectory()) {
			exist = true;
		}
		return exist;
	}
	*/
	
	public static void createTemplate() {
		final int HEADER_ROW = 0;
		final String LEVEL = "Level";
		final String  NUMBER = "Number";
		final String DESCRIPTION = "Description";
		final String QUANTITY = "Quantity";
		final String LOCATION = "Location";
		
		
		// create input directory
		File f = new File(TEMPLATE_LOCATION);
		f.mkdir();
		
		// create blank Workbook (excel file)
		Workbook wb = new XSSFWorkbook();
		
		// Create a sheet for CUSTOMER_BOM
		Sheet customerBom = wb.createSheet("CUST_BOM");
		Row customerBomHeader = customerBom.createRow(HEADER_ROW);
		customerBomHeader.createCell(0).setCellValue(LEVEL);
		customerBomHeader.createCell(1).setCellValue(NUMBER);
		customerBomHeader.createCell(2).setCellValue(DESCRIPTION);
		customerBomHeader.createCell(3).setCellValue(QUANTITY);
		customerBomHeader.createCell(4).setCellValue(LOCATION);
		
		// Create a sheet for AGILE_BOM
		Sheet agileBom = wb.createSheet("AGILE_BOM");
		Row agileBomHeader = agileBom.createRow(HEADER_ROW);
		agileBomHeader.createCell(0).setCellValue(LEVEL);
		agileBomHeader.createCell(1).setCellValue(NUMBER);
		agileBomHeader.createCell(2).setCellValue(DESCRIPTION);
		agileBomHeader.createCell(3).setCellValue(QUANTITY);
		agileBomHeader.createCell(4).setCellValue(LOCATION);
		
		String template = TEMPLATE_LOCATION + TEMPLATE_FILENAME;
		
		try(OutputStream out = new FileOutputStream(template)) {
			wb.write(out);
			out.close();
			wb.close();
			System.out.println("A new template has been created at " + template);
		} catch (Exception e) {
			System.out.println("Exception: " + e.getMessage());
		}
	}
}
