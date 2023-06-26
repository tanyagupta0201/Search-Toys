package org.example.projectoffice;

import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.edge.EdgeDriver;
import org.openqa.selenium.edge.EdgeDriverService;
import org.openqa.selenium.support.ui.Select;
import org.apache.poi.ss.usermodel.Workbook;
import com.google.common.io.Files;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import com.google.common.io.Files;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.time.Duration;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;
import java.util.concurrent.TimeUnit;
import java.io.FileOutputStream;

 

public class App {
    public static WebDriver driver = null;

    // LAUNCH DRIVER
    public static void driverSetup(String browser) {

 
        // setting up the chrome driver
        if (browser.equalsIgnoreCase("chrome")) 
        {
            // sets the property of the desired browser
            System.setProperty("webdriver.chrome.driver",
                    "C:\\\\Users\\\\DELL\\\\Downloads\\\\chromedriver_win32_new\\\\chromedriver.exe");


            // creates a new instance of the chrome driver class
            driver = new ChromeDriver();
        }


        // setting up the edge driver
        if (browser.equalsIgnoreCase("edge")) 
        {
            // sets the property of the desired browser
            System.setProperty("webdriver.edge.driver", "C:\\\\Users\\\\DELL\\\\Downloads\\\\edgedriver_win64_new\\\\msedgedriver.exe");

            // creates a new instance of the edge driver class
            driver = new EdgeDriver();
        }

        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30));        
    }

 

       // WRITTING IN EXCEL
       public static void writeExcel(String sheetName, int rNum , int cNum , int resultCount) throws IOException, InvalidFormatException 
       {
         FileInputStream fis = new FileInputStream("C:\\\\Users\\\\DELL\\\\Downloads\\\\project\\\\office\\\\projectoffice\\\\src\\\\resources\\\\excelsheet\\\\Book.xlsx");
         Workbook wb = WorkbookFactory.create(fis);
         Sheet s = wb.getSheet(sheetName);
         Row r = s.getRow(rNum);
         Cell c = r.createCell(cNum);
         c.setCellValue(resultCount);
         FileOutputStream fos = new FileOutputStream("C:\\\\Users\\\\DELL\\\\Downloads\\\\project\\\\office\\\\projectoffice\\\\src\\\\resources\\\\excelsheet\\\\Book.xlsx");
         wb.write(fos);
         fis.close();
         fos.close();
         driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30));    
      }

 

    // CLICK ADVANCED SEARCH
    public static void clickAdvancedSearch() 
    {
        // navigate to this url
        driver.get("http://www.ebay.com");

        // maximise the window
        driver.manage().window().maximize();


        WebElement advancedSearchButton = driver.findElement(By.id("gh-as-a"));
        advancedSearchButton.click();
        WebElement textBox = driver.findElement(By.id("_nkw"));

 
        textBox.sendKeys("Outdoor Toys");
        Select dropdown = new Select(driver.findElement(By.id("s0-1-17-4[0]-7[1]-_in_kw")));

 

        dropdown.selectByVisibleText("Any words, any order");


        WebElement checkBox = driver.findElement(By.id("s0-1-17-6[4]-[0]-LH_ItemCondition"));
        checkBox.click();


        WebElement location = driver.findElement(By.id("s0-1-17-6[7]-[3]-LH_PrefLoc"));
        location.click();

        WebElement search = driver.findElement(By.className("adv-form__actions")).findElement(By.tagName("button"));
        search.click();

        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30));    
    }

    public static void extractingResult() throws InvalidFormatException, IOException 
    {

        List<WebElement> allLinks = driver.findElements(By.className("s-item__link"));

        String newLink = "";
        int x = 0;
        for (WebElement link : allLinks) 
        {
            String str = link.getText();

            if (str.contains("Portable"))
            {
                newLink = (link.getAttribute("href").toString());
                x = 1;
                break;
            }
        }

        if(x == 1)
        {
            System.out.println(newLink);
            ArrayList<String> tabs = new ArrayList<String>(driver.getWindowHandles());

            driver.switchTo().window(tabs.get(0)); // switches to new tab

            driver.get(newLink);

            String str = driver.findElement(By.className("x-item-title__mainTitle")).findElement(By.tagName("span"))
                    .getText();

            System.out.println(str);
        }
        
        else 
        {
            System.out.println("item not found");

        }

        
        writeExcel("Sheet1", 1, 1, allLinks.size() - 1);
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30));    

    }

    public static void screenShot()  
    {
        File f = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
        
        try {
            Files.copy(f, new File("C:\\\\Users\\\\DELL\\\\Downloads\\\\project\\\\office\\\\projectoffice\\\\src\\\\resources\\\\screenshot\\\\screenshot1.jpg"));
        } 
        catch (IOException e) 
        {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
        
        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30));    
    }

    public static void main(String[] args) throws IOException, InvalidFormatException 
    {
        String browser;

        Scanner sc = new Scanner(System.in);
        System.out.println("Enter the browser");
        browser = sc.next();

        driverSetup(browser);
        clickAdvancedSearch();
        extractingResult();
        screenShot() ;

        driver.manage().timeouts().implicitlyWait(Duration.ofSeconds(30));    
        driver.quit();
    } 

}