import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;


import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;


import jxl.CellView;
import jxl.write.Label;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import jxl.write.Number;




public class Hello {
	private static WritableCellFormat timesBoldUnderline; 
	private static WritableCellFormat times;
	
	public static int last = 0;
	public static Map<String, Map<String,List<String>>> results = new HashMap<>();
	
	public static void main(String[] args) throws WriteException, IOException {
		System.setProperty("webdriver.chrome.driver","/Users/divakarv/Downloads/chromedriver");
		
		ChromeDriver driver=new ChromeDriver();
		
		driver.manage().window().maximize();
		driver.get("https://www.coe.annamalaiuniversity.ac.in/rgl_result.php");
		
		for(int i = 191011001; i<=191011050;i++) {
			compute(String.valueOf(i),driver);
		}
		System.out.println(results);
		driver.close();
		writeExcel("AIML");
	}
	public static void compute(String regNo, ChromeDriver driver) {
		
			
			Select se = new Select(driver.findElement(By.xpath("//*[@id=\"page-inner\"]/div/div/table/tbody/tr[1]/td/form/p[1]/select")));
			se.selectByIndex(8);
			
			try {
				
				
				WebElement input = driver.findElement(By.xpath("//*[@id=\"page-inner\"]/div/div/table/tbody/tr[1]/td/form/p[2]/input[1]"));
				input.clear();
				input.sendKeys(regNo);
				
				WebElement submit = driver.findElement(By.xpath("//*[@id=\"page-inner\"]/div/div/table/tbody/tr[1]/td/form/p[2]/input[2]"));
				submit.click();
				
				addData(regNo ,driver);
				
				
				driver.navigate().back();  
				
			} catch(Exception ex) {
				System.out.println("Exception "+regNo);
				System.out.println(ex);
				driver.navigate().back();  
			}		
	}
	public static void addData(String regNo, ChromeDriver driver){
		Map<String,List<String>> result = new HashMap<>();
		List<WebElement> subjects = driver.findElements(By.xpath("//*[@id=\"page-inner\"]/div/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr"));
	
		for(int i = 2 ; i <= subjects.size();i++) {
			List<WebElement> subjectRow = driver.findElements(By.xpath("//*[@id=\"page-inner\"]/div/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr["+i+"]/td"));
			List<String> list = new ArrayList<>();
			for(int j = 1 ; j <= subjectRow.size() ; j++) {
				WebElement subject = driver.findElement(By.xpath("//*[@id=\"page-inner\"]/div/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr["+i+"]/td["+j+"]"));
				list.add(subject.getText());
			}
			result.put(list.get(1),list);
			
		}
		if(result.isEmpty()==false) {
			results.put(regNo ,result);
		}
		
	}
	
	
	public static void writeExcel(String sheetName) throws IOException, WriteException {
		File file=new File("/Users/divakarv/Downloads/Result-DS.xls");
		WorkbookSettings settings=new WorkbookSettings();
		settings.setLocale(new Locale("en","EN"));
		WritableWorkbook workbook=Workbook.createWorkbook(file,settings);
		WritableSheet sheet = workbook.createSheet(sheetName, workbook.getNumberOfSheets());
		WritableSheet excelSheet = workbook.getSheet(workbook.getNumberOfSheets()-1);
		createLabel(excelSheet);
		createContent(excelSheet);
		workbook.write();
		workbook.close();
	}
	
	private static void createLabel(WritableSheet sheet) throws WriteException {
			WritableFont times10pt = new WritableFont(WritableFont.TIMES, 10); // Define the cell format
			times = new WritableCellFormat(times10pt); // Lets automatically wrap the cells times.setWrap(true);
			WritableFont times10ptBoldUnderline = new WritableFont( WritableFont.TIMES, 10, WritableFont.BOLD, false);
			WritableCellFormat timesBoldUnderline = new WritableCellFormat(times10ptBoldUnderline); // Lets automatically wrap the cells
			timesBoldUnderline.setWrap(true);
			CellView cv = new CellView(); 
			cv.setFormat(times); 
			cv.setFormat(timesBoldUnderline); 
			// Write a few headers
			for(String key : results.keySet()) {
				addCaption(sheet, 0, 0, "Reg No.");
				int i = 1;
				for(String courseCode : results.get(key).keySet()) {
					addCaption(sheet, i, 0, courseCode);
					i++;
				} 
				addCaption(sheet, i, 0, "CGPA");
				last = i;
				break;
			}
			 
		}
	private static void addCaption(WritableSheet sheet, int column, int row, String s) throws RowsExceededException, WriteException {
			Label label = new Label(column, row, s ); 
			sheet.addCell(label);
		}
	
	private static void createContent(WritableSheet sheet) throws WriteException, RowsExceededException {
		int row = 1;
		for(String key : results.keySet()) {
			addNumber(sheet, 0, row, Integer.parseInt(key));
			double creditPoints = 0;
			double creditHours = 0;
			boolean RA = false;
			for(String courseCode : results.get(key).keySet()) {
				int j = 1;
				while(!sheet.getCell(j,0).getContents().isEmpty()) {
					if(courseCode.equals(sheet.getCell(j,0).getContents())) {
						int grade = (int)Double.parseDouble(results.get(key).get(courseCode).get(4));
						if(grade == 0) {
							RA = true;
						}
						creditPoints += Double.parseDouble(results.get(key).get(courseCode).get(5));
						creditHours += Double.parseDouble(results.get(key).get(courseCode).get(3));
						addNumber(sheet, j, row, grade);
						break;
					}
					j++;
				}
			}
			double cgpa = 0;
			if(!RA) {
				cgpa = creditPoints / creditHours;
			}
			addNumber(sheet, last, row, Double.parseDouble(String.format("%.2f", cgpa)));
			row++;
		}
	}
	private static void addNumber(WritableSheet sheet, int column, int row, Integer integer) throws WriteException, RowsExceededException { 
		Number number = new Number(column, row, integer); 
		sheet.addCell(number);
	}
	private static void addNumber(WritableSheet sheet, int column, int row, Double d) throws WriteException, RowsExceededException { 
		Number number = new Number(column, row, d); 
		sheet.addCell(number);
	}
	private static void addLabel(WritableSheet sheet, int column, int row, String s) throws WriteException, RowsExceededException {
		Label label = new Label(column, row, s);
		sheet.addCell(label); 
	}
}
