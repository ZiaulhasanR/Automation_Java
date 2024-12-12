/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Main.java to edit this template
 */
package q1_assignment;

//import com.google.common.collect.Table.Cell;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.util.List;
import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.Month;
import java.util.Calendar;
import java.util.GregorianCalendar;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class Q1_Assignment {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws Exception {

        XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream("C:\\Users\\ziaul\\OneDrive\\Documents\\NetBeansProjects\\Q1_Assignment\\new.xlsx"));

        int day, month, year;
        GregorianCalendar gc = new GregorianCalendar();
        day = gc.get(Calendar.DAY_OF_MONTH);
        month = gc.get(Calendar.MONTH);
        year = gc.get(Calendar.YEAR);
        LocalDate myDate = LocalDate.of(year, month, day);
        DayOfWeek dayofweek = myDate.getDayOfWeek();
        String today = dayofweek.toString();
        XSSFSheet sh = wb.getSheet(today);
        if (sh == null) {
            System.out.println(today + "Does not exist");
            return;
        }
        int rowcount = sh.getLastRowNum();
        System.out.println(rowcount);
        for (int j = 1; j < rowcount; j++) {
            String data = sh.getRow(j).getCell(2).toString();
            System.out.println("Search word: " + data);

            System.setProperty("webdriver.chrome.driver", "E:\\Java\\lib\\chromedriver-win64\\chromedriver.exe");
            WebDriver driver = new ChromeDriver();
            driver.manage().window().maximize();
            Thread.sleep(2000);
            driver.get("https://www.google.com");
            //element.sendKeys("Dhaka");
            WebElement searchBox = driver.findElement(By.xpath("//*[@id=\"APjFqb\"]"));
            searchBox.sendKeys(data);
            Thread.sleep(2000);
            // searchBox.sendKeys(Keys.RETURN);

            // Fetch Google suggestions
            List<WebElement> suggestions = driver.findElements(By.className("sbct"));
            //System.out.println(suggestions.size());
            int a = Integer.MAX_VALUE, b = 0, k = 0;
            String Short = "", Long = "";
            String s = "";
            for (int i = 0; i < suggestions.size() - 1; i++) {
                s = suggestions.get(i).getText();
                k = suggestions.get(i).getText().length();
                System.out.println(k);
                if (k < a && k != 0) {
                    Short = s;
                    a = k;
                }

                if (k > b) {
                    Long = s;
                    b = k;
                }

                System.out.println(suggestions.get(i).getText());
            }
            System.out.println("Shortest String: " + Short + " (Length: " + a + ")");
            System.out.println("Longest String: " + Long + " (Length: " + b + ")");
            Thread.sleep(2000);
            //driver.quit();
            try {
                Row row = sh.getRow(j);
                if (row == null) {
                    row = sh.createRow(j);
                }

                // Write 'Long' into cell at index 2
                Cell cell = row.getCell(3);
                if (cell == null) {
                    cell = row.createCell(3); 
                }
                cell.setCellValue(Long); 

                
                Cell c = row.getCell(4);
                if (c == null) {
                    c = row.createCell(4); 
                }
                c.setCellValue(Short); 

            } catch (Exception e) {
                e.printStackTrace(); 
            }
            try (FileOutputStream fos = new FileOutputStream("C:\\Users\\ziaul\\OneDrive\\Documents\\NetBeansProjects\\Q1_Assignment\\new.xlsx")) {
                wb.write(fos);
                System.out.println("Excel file updated successfully.");
            } catch (IOException e) {
                e.printStackTrace();
            }

            // Cleanup
            

           driver.quit();
        }
        wb.close();
        //String today= getDayOfWeek(); 
    }

}
