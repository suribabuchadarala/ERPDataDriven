package driverFactory;

import java.io.File;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.testng.Reporter;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import commonFunctions.FunctionLibrary;
import config.AppUtil;
import utilities.ExcelFileUtil;

public class AppTest extends AppUtil {
String inputpath ="./FileInput/LoginData.xlsx";
String outputpath ="./FileOutput/DataDrivenResults.xlsx";
ExtentReports report;
ExtentTest logger;
@Test
public void startTest() throws Throwable
{
	//define path of html
	report = new ExtentReports("./target/Reports/DataDriven.html");
//create object for excelfile util class
	ExcelFileUtil xl = new ExcelFileUtil(inputpath);
	//count no of rows in login sheet
	int rc =xl.rowcount("Login");
	Reporter.log("No of rows are::"+rc,true);
	for(int i=1;i<=rc;i++)
	{
		logger = report.startTest("Validate Login");
		String user = xl.getcellData("Login", i, 0);
		String pass = xl.getcellData("Login", i, 1);
		//call adminloginmethod from functionlibaray class
		boolean res =FunctionLibrary.adminLogin(user, pass);
		if(res)
		{
			//write as login success into results cell
			xl.setCellData("Login", i, 2, "Login Success", outputpath);
			//write as pass into status cell
			xl.setCellData("Login", i, 3, "Pass", outputpath);
			logger.log(LogStatus.PASS, "Valid Username and Password");
		}
		else
		{
			//take scren shot and store
			File screen =((TakesScreenshot)driver).getScreenshotAs(OutputType.FILE);
			//copy screen shot into local system
			FileUtils.copyFile(screen, new File("./Screenshot/Iteration/"+i+"Loginpage.png"));
			//write as login Fail into results cell
			xl.setCellData("Login", i, 2, "Login Fail", outputpath);
			//write as Fail into status cell
			xl.setCellData("Login", i, 3, "Fail", outputpath);
			logger.log(LogStatus.FAIL, "Invalid Username and passowrd");
		}
		report.endTest(logger);
		report.flush();
	}
}
}















