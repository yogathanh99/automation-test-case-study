import io.github.bonigarcia.wdm.WebDriverManager;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.junit.After;
import org.junit.BeforeClass;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class HelloSeleniumTest {
    private WebDriver driver;

    @BeforeClass
    public static void setUpClass(){
        WebDriverManager.chromedriver().setup();
        WebDriverManager.edgedriver().setup();
    }

    @After
    public void teardown() {
        if (driver != null) {
            driver.quit();
        }
    }

    @Test
    public void firstTest() throws InterruptedException, IOException {
        driver = new ChromeDriver();
//            Create an object of File class to open xlsx file
            File file = new File("/Users/thanhvo/Class/CS493/HW4/Test.xls");

            //Create an object of FileInputStream class to read excel file
            FileInputStream inputStream = new FileInputStream(file);

            //creating workbook instance that refers to .xls file
            HSSFWorkbook wb = new HSSFWorkbook(inputStream);

            //creating a Sheet object
            HSSFSheet sheet = wb.getSheet("Sheet1");

            //get all rows in the sheet
            int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();


//            iterate over all the row to print the data present in each cell.
            for (int i = 1; i <= rowCount; i++) {
                driver.get("http://localhost:3000");
                driver.findElement(By.id("email")).sendKeys("admin@gmail.com");
                driver.findElement(By.id("password")).sendKeys("admin123");
                driver.findElement(By.cssSelector("button.input_login_button__uI0jF")).click();
                Thread.sleep(2000);
                String name=String.valueOf(sheet.getRow(i).getCell(1));
                if (name!="null")
                    driver.findElement(By.cssSelector("input.ant-input.sc-pFZIQ.fjLAgC")).sendKeys(name);
                String ID=String.valueOf(sheet.getRow(i).getCell(2));
                if (ID!="null")
                    driver.findElement(By.xpath("//input[@placeholder='Tìm kiếm theo mã nhân viên']")).sendKeys(ID);
                String birthDay=String.valueOf(sheet.getRow(i).getCell(3));
                if (birthDay!="null")
                    driver.findElement(By.xpath("//input[@placeholder='Tìm kiếm theo ngày sinh (dd/mm)']")).sendKeys(birthDay);
                driver.findElement(By.cssSelector("button.ant-btn.style_deleteButton__1-z2w.style_buttonStyle__ow-Zc > span")).click();
                Thread.sleep(2000);
                Integer displayData =driver.findElements(By.cssSelector("ul.ant-pagination")).size();
                Integer displayEmpty= driver.findElements(By.cssSelector("div.ant-empty-image")).size();

                Row row = sheet.getRow(i);
                Cell cellAnswer = row.createCell(4);
                cellAnswer = row.getCell(4);
                if (displayData >0) {
                    cellAnswer.setCellValue("PASS-Has Data");
                }
                if (displayEmpty>0){
                    cellAnswer.setCellValue("PASS-Empty Data");
                }

                FileOutputStream fileOutputStream = new FileOutputStream(file);
                wb.write(fileOutputStream);
                fileOutputStream.close();

                driver.manage().deleteAllCookies();
                Thread.sleep(3000);
            }
            inputStream.close();
            driver.quit();
    }
}
