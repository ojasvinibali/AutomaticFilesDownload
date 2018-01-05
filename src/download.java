import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.TimeUnit;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;


public class download {
	public static void main(String[] args) throws IOException {

		int flag = 0;
		String studentlabelxpath = "//*[contains(@id,'ext-comp-1031') and contains(@type,'text')]";
		String downloadxpath = "//a[contains(@title,'.pdf')]";

		System.setProperty("webdriver.chrome.driver", "C:\\Users\\o1\\Downloads\\Survey1\\Survey1\\chromedriver_win32\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		//WebDriver driver = new FirefoxDriver();
		driver.get("https://uncc.starfishsolutions.com/starfish-ops/support/login.html");
		//driver.get("https://uncc-test.starfishsolutions.com/starfish-prod/support/login.html");

		//Login
		driver.findElement(By.xpath("//*[@id=\"username\"]")).sendKeys("jchoudh1");
		driver.findElement(By.xpath("//*[@id=\"password\"]")).sendKeys("May_Dec");
		driver.findElement(By.xpath("//*[@id=\"shibboleth-login-button\"]")).click();


		driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);

		//Click students
		driver.findElement(By.xpath("//*[@id='topNavStudentsButton']")).click(); // click students
		driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);


		String inputFile = "C:\\Users\\o1\\Downloads\\Survey1\\Survey1\\src\\download1.xls";
		File inputWorkbook = new File(inputFile);
		Workbook w;

		try {
			w = Workbook.getWorkbook(inputWorkbook);
			// Get the first sheet
			Sheet sheet = w.getSheet(0);

			for (int j = 1; j < sheet.getRows(); j++) {
				ArrayList<String> inputDataArray = new ArrayList<String>(25);
				for (int i = 0; i < sheet.getColumns(); i++) {
					Cell cell = sheet.getCell(i, j);
					inputDataArray.add(cell.getContents());
					//System.out.println(inputDataArray.get(i));
				}

				{
					String name = inputDataArray.get(0).trim();
					String namexpath = "//li[contains(text(),'" + name.toLowerCase() + "')]";
					driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
					driver.findElement(By.xpath(studentlabelxpath)).sendKeys(name);

					//Wait till Value appears in drop down
					WebDriverWait wait = new WebDriverWait(driver, 60);
					String notfound = "//div[contains(text(),'No people found')]";

					//driver.manage().timeouts().implicitlyWait(30, TimeUnit.SECONDS);
					if(driver.findElement(By.xpath(notfound)).isDisplayed()){
						System.out.println("Not Found : " + name);
					}
					else {
						 wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(namexpath)));
			
						//Click on Result
						driver.findElement(By.xpath(namexpath)).click();

						String home = System.getProperty("user.home");
						String mypath = home + File.separator + "Downloads" + File.separator + name;

						//create directory
						File dir = new File(mypath);
						dir.mkdir();


						driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);

						List<WebElement> a = driver.findElements(By.xpath("//a[contains(@title,'.pdf') or contains(@title,'.tif')]"));
						int length = a.size();
						if (length <= 0) {
							System.out.println("Not Downloaded : " + name);
						} else {
							for (int i = 0; i < length; i++) {
								//((JavascriptExecutor)driver).executeScript("argument[0].click();",);
								JavascriptExecutor executor = (JavascriptExecutor) driver;
								executor.executeScript("arguments[0].click();", a.get(i));
							}

							Thread.sleep(10000);
							File folder = new File(home + File.separator + "Downloads");
							File[] listOfFiles = folder.listFiles();

							for (int i = 0; i < listOfFiles.length; i++) {

								if (listOfFiles[i].isFile() && listOfFiles[i].getName().endsWith(".pdf")) {

									File f = new File(home + File.separator + "Downloads" + File.separator + listOfFiles[i].getName());

									///f.renameTo(new File(mypath+File.separator + name + i + ".pdf"));
									f.renameTo(new File(mypath + File.separator + listOfFiles[i].getName()));

									//System.out.println(listOfFiles[i].getName());
								} else if (listOfFiles[i].isFile() && listOfFiles[i].getName().endsWith(".tif")) {
									File f = new File(home + File.separator + "Downloads" + File.separator + listOfFiles[i].getName());

									f.renameTo(new File(mypath + File.separator + listOfFiles[i].getName()));

									//System.out.println(listOfFiles[i].getName());

								} else if (listOfFiles[i].isFile() && listOfFiles[i].getName().endsWith(".tiff")) {
									File f = new File(home + File.separator + "Downloads" + File.separator + listOfFiles[i].getName());

									f.renameTo(new File(mypath + File.separator + listOfFiles[i].getName()));

									//System.out.println(listOfFiles[i].getName());

								}
							}

							System.out.println("conversion is done : " + name);
						}
						// System.out.println("Done  :" + name);
					}
					// }


					driver.navigate().refresh();
					driver.manage().timeouts().implicitlyWait(50, TimeUnit.SECONDS);

				}


			}
		}catch (BiffException e)
		{
			e.printStackTrace();
		} catch (InterruptedException e) {
			e.printStackTrace();
		}


	}

}