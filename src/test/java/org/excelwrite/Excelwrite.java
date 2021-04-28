package org.excelwrite;



import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Name;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.support.ui.Select;

public class Excelwrite {
	public static void main(String[] args) throws IOException {
		System.setProperty("webdriver.edge.driver","C:\\Users\\tprab\\eclipse-workspace\\Day3ExcelWriteandupdate\\Driver\\msedgedriver.exe");
		WebDriver driver = new EdgeDriver();
		driver.get("http://demo.automationtesting.in/Register.html");
		
		WebElement country = driver.findElement(By.id("countries"));
		Select sl = new Select(country);
		List<WebElement> list = sl.getOptions();
		
		File file = new File("C:\\Users\\tprab\\eclipse-workspace\\Day3ExcelWriteandupdate\\excelsheets\\countries.xlsx");
		
		Workbook wr = new XSSFWorkbook();
		Sheet sh = wr.createSheet();
		
		for (int i =0 ; i <=list.size()-1; i++) {
			WebElement element = list.get(i);
			String list1 = element.getText();
			System.out.println(list1);
			
			Row r = sh.createRow(i);
			Cell c = r.createCell(0);
			c.setCellValue(list1);
				}
		
		FileOutputStream out = new FileOutputStream(file);
		wr.write(out);
		System.out.println("done");
		
		System.out.println("sanjai");
		System.out.println("rajkumar");
		
			
			
			
		}
		
		
		
		
		
		
		
		
	}	

