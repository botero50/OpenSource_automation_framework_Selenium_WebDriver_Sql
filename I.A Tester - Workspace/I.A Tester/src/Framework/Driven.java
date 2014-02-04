package Framework;
import java.sql.*;
import java.awt.Rectangle;
import java.awt.Robot;
import java.awt.Toolkit;
import java.awt.image.BufferedImage;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;

import javax.imageio.ImageIO;
import com.microsoft.sqlserver.jdbc.*;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.internal.annotations.TestAnnotation;
import java.util.Date;

public class Driven extends TestAnnotation
{
	//WorkSpace
	public String WorkspaceDirectory="";
	//Connection
	//Database Conection
	public Connection ConectionDriver=null;
	//Steeps
	public ResultSet RsScript= null;
	//Data
	public ResultSet Rscases= null;
	//Execition Scripts
	public ResultSet RsExecutions= null;
	
	//Interface Configuration
	public int InterfaceId=0;
	public String InterfaceName=null;
	
	//Driver Configurations
    public WebDriver driver;
    //Automation Object
    public WebElement AutomationObject= null;
    //Variable of control
    public boolean Validation = true;
    
    //Enumerate Cases
    public enum Type 
	{
    	operate, validate, open, print, typekeys,navigate, error;
		public static Type getValue(String str)
	    {
	        try {return valueOf(str);}
	        catch (Exception ex) {return error;}
	    }
	}
    public enum OperateType
    {
		check,clear,click,select,set,type,uncheck,error;
		public static OperateType getValue(String str)
		{
		    try {return valueOf(str);}
		    catch (Exception ex) {return error;}
		}
    }
    public enum ValidateType
    {
		cheacked,enabled,text,exist,len,visible,unchecked,error;
		public static ValidateType getValue(String str)
		{
		    try {return valueOf(str);}
		    catch (Exception ex) {return error;}
		}
    }
    public enum Navigator 
	{
    	ie, ff, chr, error;
		public static Navigator getValue(String str)
	    {
	        try {return valueOf(str);}
	        catch (Exception ex) {return error;}
	    }
	}
    public enum ObjectIdType 
	{ 
    	id,name,xpath,classname ,tagname,linktext,cssselector , error;
		public static ObjectIdType getValue(String str)
	    {
	        try {return valueOf(str);}
	        catch (Exception ex) {return error;}
	    }
	}
    public void SetUpVars(String WorkSpaceProyect)
    {
    	WorkspaceDirectory=WorkSpaceProyect;
    }
    
    public boolean Connection(String ServerName,String DatabaseName,String User, String Password)
    {
    	try {
			Class.forName("com.microsoft.sqlserver.jdbc.SQLServerDriver");
		} catch (ClassNotFoundException e1) {e1.printStackTrace();}
    	
    	SQLServerDataSource ds = new SQLServerDataSource();
        ds.setServerName(ServerName);
        ds.setDatabaseName(DatabaseName);         
        ds.setUser(User);
        ds.setPassword(Password);
                       	
        try {ConectionDriver = ds.getConnection(); return true;} 
        catch (SQLException e) {return false;}
    }
    public void OpenExecutions()
    {
    	try {
			RsExecutions = ConectionDriver.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY).executeQuery("SELECT Ex.InterfaceId, Inter.InterfaceSciptName FROM [6TbExecutor] Ex inner join [1TbInterface] Inter on Ex.InterfaceId = Inter.InterfaceId  where [ExecutionStatus] ='False' and GETDATE() > [ExecutionDate]  ORDER BY [ExecutionDate]");
		} catch (SQLException e) {}
    }
    public void OpenResulsets(Integer IdInterface, String InterfaceName)
    {
    	try
    	{
    		RsScript = ConectionDriver.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY).executeQuery("SELECT * FROM [3TbScript] where [Implement] ='True' and InterfaceId = " + Integer.toString(IdInterface) + "ORDER BY [ImplementOrder]");
			Rscases = ConectionDriver.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY).executeQuery("SELECT * FROM [TC" +  InterfaceName + "] ORDER BY [IdCase]");
			//Clear database
			ConectionDriver.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY).executeQuery("DELETE FROM [5TbReports] where InterfaceId = " +  Integer.toString(IdInterface));
    	} catch (SQLException e) {}
    }
    public void UpdateExecution(Integer IdInterface)
    {
    	try
    	{
			ConectionDriver.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY).executeQuery("update [6TbExecutor] set [ExecutionStatus]='true' where InterfaceId = " +  Integer.toString(IdInterface));
    	} catch (SQLException e) {}
    }
    public void CloseResulsets()
    {
    		RsScript=null; 
			Rscases=null;
			driver.close();
    }
    
    public void run() throws SQLException	
    {
    	String SteepType = "setup";
    	while ( Rscases.next()) 
        {
    		//Cases Vars
    		String StrSpectedResult =NullControl(Rscases.getString("ExpectedResult"));
    		Integer StrCaseId =Rscases.getInt("IdCase");
    		
    		Validation=true;
    		
    		while (RsScript.next())
    		{
    			Boolean ExecuteCase = true;
    			if(!SteepType.equalsIgnoreCase("setup"))
				{
    				if (RsScript.getString("SteepType").trim().equalsIgnoreCase("setup"))
    					ExecuteCase = false;
				}
    			
    			if (ExecuteCase)
    			{
	    			//Script Vars
	    			Integer StrSteepId = RsScript.getInt("SteepId");
	    			String StrOperateType = GetString(NullControl(RsScript.getString("OperateType")));
	    			String StrOperateValue = GetString(NullControl(RsScript.getString("OperateValue")));
	    			String StrOptionalData1 = GetString(NullControl(RsScript.getString("OptionalData1")));
	    			String StrOptionalData2 = GetString(NullControl(RsScript.getString("OptionalData2")));
	    			
	    			if (!(StrOperateType.equalsIgnoreCase("|jump|") && StrOperateValue.equalsIgnoreCase("|jump|")))
	    			{
		    			switch (Type.getValue(RsScript.getString("ImplementType").toLowerCase()))	
		    			{
		    			case operate:
		    				Validation = GetObject(GetString(RsScript.getString("AutomationObject")));
		    				if(Validation) 
		    					Validation = Operate(StrOperateType, StrOperateValue, StrOptionalData1, StrOptionalData2);
							if(!Validation)
								ReportEvent(false, "Error trying to validate ( " + StrSpectedResult + " ) In Operate Keyword",StrCaseId,StrSteepId);
							break;
						case validate:		
							Validation = GetObject(GetString(RsScript.getString("AutomationObject")));
							if(Validation) 
								Validation = Validate(StrOperateType, StrOperateValue, StrOptionalData1, StrOptionalData2);
							if(!Validation)
								ReportEvent(false,"Error trying to validate ( " + StrSpectedResult + " )  In Validate Keyword",StrCaseId,StrSteepId);
							break;
						case print:	
							Validation = ReportEvent(true, StrSpectedResult,StrCaseId,StrSteepId);
							if(!Validation)
								ReportEvent(false,"Error trying to validate ( " + StrSpectedResult + " )  In Operate Keyword",StrCaseId,StrSteepId);
							break;
						case open:	
							try
							{
								switch (Navigator.getValue(RsScript.getString("AutomationObject").toLowerCase()))	
				    			{
									case ie:
										System.setProperty("webdriver.ie.driver","IEDriverServer.exe");
								    	driver = new InternetExplorerDriver();
								        driver.get(GetString(RsScript.getString("OperateValue").toLowerCase())); 
								        driver.manage().window().maximize();
								        Validation=true;
										break;
									case ff:
								    	driver = new FirefoxDriver();
								        driver.get(GetString(RsScript.getString("OperateValue").toLowerCase())); 
								        driver.manage().window().maximize();
								        Validation=true;
										break;
									case chr:
								    	driver = new ChromeDriver();
								        driver.get(GetString(RsScript.getString("OperateValue").toLowerCase())); 
								        driver.manage().window().maximize();
								        Validation=true;
										break;
								default:
									Validation =false;
									ReportEvent(false,"Error trying to validate ( " + StrSpectedResult + " )  Unrecognized Browser",StrCaseId,StrSteepId);
									break;
				    			}
			    			}catch(Exception ex)
							{
								Validation=false;
								ReportEvent(false,"Error trying to validate ( " + StrSpectedResult + " )  In TypeKeys Keyword",StrCaseId,StrSteepId);
							}
							break;
						case navigate:
							try
							{
							 driver.get(GetString(RsScript.getString("OperateValue").toLowerCase()));
							}catch(Exception ex)
							{
								Validation=false;
								ReportEvent(false,"Error trying to validate ( " + StrSpectedResult + " )  In TypeKeys Keyword",StrCaseId,StrSteepId);
							}
							 break;
						case typekeys:
							Validation = Typekeys(GetString(StrOperateValue));
							if(!Validation)
								ReportEvent(false,"Error trying to validate ( " + StrSpectedResult + " )  In TypeKeys Keyword",StrCaseId,StrSteepId);
							break;
						default:
							ReportEvent(false,"Error trying to validate ( " + StrSpectedResult + " ) Unrecognized Keyword",StrCaseId,StrSteepId);
							Validation = false;
							break;
		    			}
		    			
		    			if(!Validation)
		    				break;
		    		}
	    		}
    		}
    		RsScript.first();
    		RsScript.previous();
    		SteepType = "Steep";
        }
    }
    public String GetString(String StrData)
    {
    	try {
	    	if(StrData.substring(0, 3).equalsIgnoreCase("om_"))
	    	{  		
	    		ResultSet RsDato = ConectionDriver.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY).executeQuery("SELECT [IdObjectType] + ';' + [ObjectIdentificationType] + ';' + [ObjectIdentificator] as Dato FROM [2TbObjectMap] where LOWER ([ObjectName]) ='" + StrData.substring(3, StrData.length()).toLowerCase().toString() +"' And InterfaceId =" + Integer.toString(InterfaceId));
	    		RsDato.next();
	    		return RsDato.getString("Dato");
	    	}
	    	else if(StrData.substring(0, 3).equalsIgnoreCase("od_"))
	    	{
	    		ResultSet rs_Get=Rscases;
				return rs_Get.getString(StrData.substring(3, StrData.length()));
	    	}else
	    	{
	    		return StrData;
	    	}
		} catch (Exception e) 
		{
			return StrData;
		}
    }
    public String NullControl(String StrData)
    {
    	if(StrData != null)
    	{
    		return StrData;
    	}else
    	{
    		return "";
    	}
    }
    public boolean Operate(String StrOperateType,String StrOperateValue,String StrOptionalData1,String StrOptionalData2)
    {
    	try
    	{
	    	switch (OperateType.getValue(StrOperateType.toLowerCase()))	
			{
				case set:
					AutomationObject.sendKeys(StrOperateValue);
					return true;
				case check:
					if (!AutomationObject.isSelected())	AutomationObject.click();
					return true;
				case uncheck:
					if (AutomationObject.isSelected())	AutomationObject.click();
		    	     return true;
				case clear:
					AutomationObject.sendKeys("");
					return true;
				case click:
					AutomationObject.click();
					return true;
				case select:
					new Select(AutomationObject).selectByVisibleText(StrOperateValue);
					return true;
				case type:
					new Actions(driver).moveToElement(AutomationObject).perform();
					Typekeys(StrOperateValue);
					return true;
				default:
					return false;
				}
    	}
    	catch(Exception Ex)
    	{
    		return false;
    	}
    }
    public boolean Validate(String StrOperateType,String StrOperateValue,String StrOptionalData1,String StrOptionalData2)
    {
    	try
    	{
	    	switch (ValidateType.getValue(StrOperateType.toLowerCase()))	
			{
				case cheacked:
					if (BoolCompare (AutomationObject.isSelected(), Boolean.parseBoolean(StrOperateValue)))
						return true;
					else
						return false;
				case unchecked:
					if (BoolCompare (!AutomationObject.isSelected(), Boolean.parseBoolean(StrOperateValue)))
						return true;
					else
						return false;
				case exist:
					if (BoolCompare(ValidateExistence(driver,GetString(RsScript.getString("AutomationObject")).split(";")[1],GetString(RsScript.getString("AutomationObject")).split(";")[2],05), Boolean.parseBoolean(StrOperateValue)))
						return true;
					else
						return false;
				case enabled:
					if (BoolCompare (!AutomationObject.isEnabled(), Boolean.parseBoolean(StrOperateValue)))
						return true;
					else
						return false;
				case len:
					if (AutomationObject.getText().trim().length() == StrOperateValue.length())
						return true;
					else
						return false;
				case text:
					if(AutomationObject.getText().trim().contentEquals(StrOperateValue))
						return true;
					else
						return false;
				case visible:
					if (BoolCompare (!AutomationObject.isDisplayed(), Boolean.parseBoolean(StrOperateValue)))
						return true;
					else
						return false;
				default:
					return false;
			}
		}
		catch(Exception Ex)
		{
			return false;
		}
    }
    public boolean GetObject(String Object)
    {
    	try
    	{
    		if(ValidateExistence(driver,Object.split(";")[1],Object.split(";")[2],05))
    		{
		    	switch (ObjectIdType.getValue(Object.split(";")[1]))	
				{
					case id:
						AutomationObject = driver.findElement(By.id(Object.split(";")[2]));
						return true;
					case name:
						AutomationObject = driver.findElement(By.name(Object.split(";")[2]));
						return true;
					case classname:
						AutomationObject = driver.findElement(By.className(Object.split(";")[2]));
						return true;
					case cssselector:
						AutomationObject = driver.findElement(By.cssSelector(Object.split(";")[2]));
						return true;
					case linktext:
						AutomationObject = driver.findElement(By.linkText(Object.split(";")[2]));
						return true;
					case tagname:
						AutomationObject = driver.findElement(By.tagName(Object.split(";")[2]));
						return true;
					case xpath:
						AutomationObject = driver.findElement(By.xpath(Object.split(";")[2]));
						return true;	
					default:
						return false;
					}
    			}    		
    			else
        			return false;
			}
			catch(Exception Ex)
			{
				return false;
			}
    }
    private static boolean ValidateExistence(WebDriver driver, String Type, String Id ,int Time) throws InterruptedException
	{
		try
		{
			WebDriverWait wait = new WebDriverWait(driver, Time);
			if(Type.equalsIgnoreCase("name"))
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.name(Id)));
			else if(Type.equalsIgnoreCase("id"))
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.id(Id)));
			else if(Type.equalsIgnoreCase("linktext"))
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.linkText(Id)));
			else if(Type.equalsIgnoreCase("classname"))
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.className(Id)));
			else if(Type.equalsIgnoreCase("cssselector"))
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.cssSelector(Id)));
			else if(Type.equalsIgnoreCase("linktext"))
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.linkText(Id)));
			else if(Type.equalsIgnoreCase("tagname"))
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.tagName(Id)));
			else if(Type.equalsIgnoreCase("xpath"))
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(Id)));
			else
				wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath(Id)));
		}catch(Exception ex)
		{
			return false;
		}		
		return true;
	}
    public Boolean BoolCompare(Boolean ActualData, Boolean Expected)
    {
    	if(ActualData == Expected)
    	{
    		return true;
    	}else
    	{
    		return false;
    	}
    }
    public boolean Typekeys(String text) {
       try {
           Robot robot = new Robot();
           for(int i=0;i<text.length();i++) {
               robot.keyPress(text.charAt(i));
           }
       } catch(java.awt.AWTException exc) {
    	   return false;
       }
       return true;
    }
    public Boolean ReportEvent(boolean Type, String Report,Integer CaseId, Integer Steep)
    {
    	try
    	{
    		Date date = new Date();
    		String ImageName= "\\1. Html\\1. Images\\" + date.toString().replace(":", "") + ".png";
    		BufferedImage imageRobot = new Robot().createScreenCapture(new Rectangle(Toolkit.getDefaultToolkit().getScreenSize()));
            ImageIO.write(imageRobot, "png", new File(WorkspaceDirectory + ImageName));

	    	File f = new File(WorkspaceDirectory + ImageName);
	        FileInputStream in = new FileInputStream(f);
	        byte[] image = new byte[(int) f.length()];
	        in.read(image);
	        String sql = "INSERT INTO [5TbReports] ([InterfaceId],[ReportType],[DataCaseId],[IdSteep],[ReportResult],[ReportEvidence],[ImageEvidence])VALUES(?,?,?,?,?,?,?)";
	        PreparedStatement stmt = ConectionDriver.prepareStatement(sql);
	        stmt.setInt(1, InterfaceId);
	        stmt.setBoolean(2, Type);
	        stmt.setInt(3, CaseId);
	        stmt.setInt(4, Steep);
	        stmt.setString(5, Report);
	        stmt.setString(6, ImageName);
	        stmt.setBytes(7, image);
	        stmt.executeUpdate();
	        stmt.close();
	        in.close();
	        return true;
    	}catch(Exception Ex)
    	{
    		return false;
    	}
    }
    public void GenerateReport(Integer StrInterfaceId, String StrInterfaceName) throws SQLException	
    {
			Rscases.first();
			Rscases.previous();
    	
			try 
			{
				String ln = System.getProperty("line.separator");
				File outFile = new File(WorkspaceDirectory + "\\1. Html\\" + InterfaceName + ".html");
				BufferedWriter writer = new BufferedWriter(new FileWriter(outFile));
				Date date = new Date();

				writer.write("<html>" +
						"<title> I.A Tester </title> " +
												
						"<SCRIPT language=\"JavaScript\" type=\"text/javascript\"> " + ln +
						"var newwindow = '' "  + ln +
						"function popitup(url) { "  + ln +
						"if (newwindow.location && !newwindow.closed) { "  + ln +
						  "  newwindow.location.href = url; "  + ln +
						 "   newwindow.focus(); } "  + ln +
						" else { "  + ln +
							 " newwindow=window.open(url,'htmlname','width=1100,height=688,resizable=1'); "  + ln +
						    " newwindow.moveTo(0, 0); "  + ln +
						    "      } "  + ln +
						"} </SCRIPT>"  + ln +
						// Based on JavaScript provided by Peter Curtis at www.pcurtis.com			
						
						"<body><TABLE BORDER=1 align=center>" +
						"<TR><TH colspan=2>Report Execution</TH></TR>" +
						
						"<TR><TH width=200>Interface Name</TH>" +
						"<TD width=200>" +  StrInterfaceName + "</TD></TR><TR>" +
						
						"<TH>Execution Date</TH>   " +
						"<TD>" +   date.toString() + "</TD></TR></TABLE>" +
						"<TABLE BORDER=1 width=100%> " + ln);

    
    	while ( Rscases.next()) 
        {
    		writer.write("<TR><TH colspan=3>Datacase " + Rscases.getString("IdCase") +"</TH> </TR>");
    		writer.write(
    			"<TR>" +
    			"<TH colspan=1 width=350>Description</TH> " +
    			"<TH colspan=1 width=710>Steeps</TH> " +
    			"<TH colspan=1 width=400>Result</TH> " +
    			"</TR>"  + ln);
    		
    		String Steeps=
    				"Select Sc.[ImplementOrder],Sc.[SteepId],Sc.[ImplementType],Sc.[AutomationObject],Sc.[OperateType],Sc.[OperateValue],'True' as [ReportType],'' as [ReportResult],'' as [ReportEvidence] from [3TbScript] Sc " +
    				" left join [5TbReports] Re on Sc.InterfaceId = Re.InterfaceId and Sc.SteepId != Re.IdSteep" +
    				" where Sc.InterfaceId = " +  Integer.toString(StrInterfaceId)  + " and Re.DataCaseId = " + Rscases.getString("IdCase") +
    				" union all" +
    				" Select Sc.[ImplementOrder],Sc.[SteepId],Sc.[ImplementType],Sc.[AutomationObject],Sc.[OperateType],Sc.[OperateValue],Re.[ReportType],Re.[ReportResult],Re.[ReportEvidence] from [3TbScript] Sc" + 
    				" left join [5TbReports] Re on Sc.InterfaceId = Re.InterfaceId and Sc.SteepId = Re.IdSteep" +
    				" where Sc.InterfaceId = " +  Integer.toString(StrInterfaceId)  + " and Re.DataCaseId = " + Rscases.getString("IdCase") +
    				" order by Sc.ImplementOrder";
    		ResultSet RsSteep = ConectionDriver.createStatement(ResultSet.TYPE_SCROLL_INSENSITIVE, ResultSet.CONCUR_READ_ONLY).executeQuery(Steeps);;
    		
	    		writer.write("<TD align=center>" +  Rscases.getString("DescriptionCase") +  "</TD>"  + ln);
		    		   
	    		writer.write("<TD align=center><table border =1>" +
	    				"<TH>Steep Id</TH>" +
	    				"<TH>Result</TH>" +
	    				"<TH>Implement Type</TH>" +
	    				"<TH>Automation Object</TH>" +
	    				"<TH>Operate Type</TH>" +
	    				"<TH>Operate Value</TH>" +
	    				"<TH>Image</TH>" + ln);
	    		String Color ="bgcolor=\"#2EFE64\"";
	    		String ExpectedResult = Rscases.getString("ExpectedResult");
	    	while (RsSteep.next())
	    	{				
	    		String ReportType = RsSteep.getString("ReportType");
				if(ReportType.contentEquals("1") && !Color.contentEquals("bgcolor=\"#FF0000\""))
				{
					ReportType = "Pass";
				}
				else
				{
					if(!Color.contentEquals("bgcolor=\"#FF0000\""))ExpectedResult= RsSteep.getString("ReportResult");
					ReportType = "Fail";
					Color = "bgcolor=\"#FF0000\"";
				}
				writer.write("<tr " + Color + "><td align=center>" + RsSteep.getString("SteepId") + "</td>" + ln);
				writer.write("<td align=center>" + ReportType + "</td>" + ln);
				writer.write("<td align=center>" + RsSteep.getString("ImplementType") + "</td>" + ln);
				writer.write("<td align=center>" + GetString(NullControl(RsSteep.getString("AutomationObject"))) + "</td>" + ln);
				writer.write("<td align=center>" + GetString(NullControl(RsSteep.getString("OperateType"))) + "</td>" + ln);
				writer.write("<td align=center>" + GetString(NullControl(RsSteep.getString("OperateValue"))) + "</td>" + ln);
				
				String Evidencia=GetString(NullControl(RsSteep.getString("ReportEvidence")));
				if (Evidencia.contentEquals(""))
					writer.write("<td align=center></td></tr>" + ln);
				else
					writer.write("<td align=center> <a href= \"javascript:popitup('.." + Evidencia.replace('\\', '/') + "')\" target=\"blank\">Image</a></td></tr>" + ln);
				
				
    		}
    		writer.write("</table></TD><TD align=center>" + ExpectedResult + "</TD>" + ln); 
	    	writer.write("</TR>");
        }
    	writer.write("</TABLE></body></html>"); 
    	writer.close();
    	} catch (IOException e) {}
    }
}




