import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.Alert;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.remote.CapabilityType;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.Select;

public class BSNLTest1 {

	public static void main(String[] args) throws InterruptedException, IOException {
		// TODO Auto-generated method stub
		
		   //DesiredCapabilities cap= new DesiredCapabilities();
		    
		   //cap.setCapability(CapabilityType.ACCEPT_SSL_CERTS, true);
		
		   System.setProperty("webdriver.chrome.driver", "C:\\Users\\Public\\chromedriver.exe");
		 
		    
		    WebDriver driver= new ChromeDriver();
		    		    		    		    
		    //driver.get("http://ktkint.bsnl.co.in/ktkint");
		    
		    		    
		    driver.get("https://ktkint.bsnl.co.in/ktkint/cdr_reports/");
		    		    
		    Thread.sleep(40000);
		    
		    driver.manage().window().maximize();
		    
		   //Alert alert= driver.switchTo().alert();
		   
	       driver.findElement(By.linkText("MIS Report")).click();
		    
		    Thread.sleep(5000);
		    
		    Select SSA=new Select(driver.findElement(By.id("ssa_name")));
		    
		    SSA.selectByValue("BIDAR");
		    
		    Thread.sleep(5000);
		    
             Select month=new Select(driver.findElement(By.id("month")));
		    
             month.selectByValue("06-Jun");
             
             Thread.sleep(5000);
             
             Select year=new Select(driver.findElement(By.id("year")));
 		    
             year.selectByValue("2021");
             
             Thread.sleep(5000);
             
             //driver.findElement(By.xpath("//input[type='Submit']")).click();;
                          
              driver.findElement(By.name("Submit")).click();
		    
              //driver.findElement(By.cssSelector(cssSelector)
              
              WebElement table = driver.findElement(By.xpath("/html/body/table/tbody")); 	
              
              int rows= table.findElements(By.xpath("tr")).size();
             
             System.out.println(rows);
             
             for (int r=1;r<=rows; r++) {
            	 
            	 FileInputStream fis = new FileInputStream("C:\\Users\\Public\\Test.xlsx");
   	    	     XSSFWorkbook workbook = new XSSFWorkbook(fis);
   	    	     XSSFSheet sheet = workbook.getSheet("Sheet1");
            	 
            	 String x1= table.findElement(By.xpath("tr["+r+"]/td[1]")).getText();
            	 
            	    Row row = sheet.createRow(r-1);
       	    	    org.apache.poi.ss.usermodel.Cell cell= row.createCell(0);
       	    	    cell.setCellValue(x1);
            	 
            	 String x2= table.findElement(By.xpath("tr["+r+"]/td[2]")).getText();
            	             	    
     	    	    cell= row.createCell(1);
     	    	    cell.setCellValue(x2);
            	 
            	 String x3= table.findElement(By.xpath("tr["+r+"]/td[3]")).getText();
            	 
  	    	    cell= row.createCell(2);
  	    	    cell.setCellValue(x3);
            	 
            	 String x4= table.findElement(By.xpath("tr["+r+"]/td[4]")).getText();
            	 
            	    cell= row.createCell(3);
      	    	    cell.setCellValue(x4);
            	 
            	 String x5= table.findElement(By.xpath("tr["+r+"]/td[5]")).getText();
            	 
         	    cell= row.createCell(4);
  	    	    cell.setCellValue(x5);
            	 
            	 String x6= table.findElement(By.xpath("tr["+r+"]/td[6]")).getText();
            	 
            	 cell= row.createCell(5);
   	    	    cell.setCellValue(x6);
            	 
            	 String x7= table.findElement(By.xpath("tr["+r+"]/td[7]")).getText();
            	 
            	 cell= row.createCell(6);
    	    	    cell.setCellValue(x7);
             	 
            	 
            	 String x8= table.findElement(By.xpath("tr["+r+"]/td[8]")).getText();
            	 
            	 cell= row.createCell(7);
    	    	    cell.setCellValue(x8);
             	 
    	    	    
            	 String x9= table.findElement(By.xpath("tr["+r+"]/td[9]")).getText();
            	 
            	 cell= row.createCell(8);
    	    	    cell.setCellValue(x9);
             	 
    	    	    
            	 String x10= table.findElement(By.xpath("tr["+r+"]/td[10]")).getText();
            	 
            	 cell= row.createCell(9);
    	    	    cell.setCellValue(x10);
             	 
            	 String x11= table.findElement(By.xpath("tr["+r+"]/td[11]")).getText();
            	 
            	 cell= row.createCell(10);
    	    	    cell.setCellValue(x11);
             	 
            	 String x12= table.findElement(By.xpath("tr["+r+"]/td[12]")).getText();
            	 
            	 cell= row.createCell(11);
    	    	    cell.setCellValue(x12);
             	 
            	 String x13= table.findElement(By.xpath("tr["+r+"]/td[13]")).getText();
            	 
            	 cell= row.createCell(12);
    	    	    cell.setCellValue(x13);
             	 
            	 String x14= table.findElement(By.xpath("tr["+r+"]/td[14]")).getText();
            	 
            	 cell= row.createCell(13);
    	    	    cell.setCellValue(x14);
             	 
            	 String x15= table.findElement(By.xpath("tr["+r+"]/td[15]")).getText();
            	 
            	 cell= row.createCell(14);
    	    	    cell.setCellValue(x15);
             	             	 
            	 String x16= table.findElement(By.xpath("tr["+r+"]/td[16]")).getText();
            	 
            	 cell= row.createCell(15);
    	    	    cell.setCellValue(x16);
             	 
            	 
            	 String x17= table.findElement(By.xpath("tr["+r+"]/td[17]")).getText();
            	 
            	 cell= row.createCell(16);
    	    	    cell.setCellValue(x17);
             	 
    	    	 String x18= table.findElement(By.xpath("tr["+r+"]/td[18]")).getText();
    	    	 
    	    	 cell= row.createCell(17);
    	    	    cell.setCellValue(x18);
             	 
            	 String x19= table.findElement(By.xpath("tr["+r+"]/td[17]")).getText();
            	 
            	 cell= row.createCell(18);
    	    	    cell.setCellValue(x19);
             	 
            	 String x20= table.findElement(By.xpath("tr["+r+"]/td[20]")).getText();
            	 
            	 cell= row.createCell(19);
    	    	    cell.setCellValue(x20);
             	 
            	 String x21= table.findElement(By.xpath("tr["+r+"]/td[21]")).getText();
            	 
            	 cell= row.createCell(20);
    	    	    cell.setCellValue(x21);
             	 
            	 String x22= table.findElement(By.xpath("tr["+r+"]/td[22]")).getText();
            	 
            	 cell= row.createCell(21);
    	    	    cell.setCellValue(x22);
             	           	             	    	    	
   	    	     	       	    	  
   	    	    FileOutputStream fos = new FileOutputStream("C:\\Users\\Public\\Test.xlsx");
   	    	    workbook.write(fos);
   	    	    fos.close();
   	    	               
                  	 
             }
             
             System.out.println("END OF WRITING DATA IN EXCEL");
             
             driver.close();
              
	}

}
