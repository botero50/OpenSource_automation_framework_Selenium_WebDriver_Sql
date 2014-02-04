package Framework;

import java.sql.ResultSet;
import java.sql.SQLException;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;

public class Executor 
{
	Driven DrivenProcess = new Driven();
	@BeforeTest
	public void SetUp()
	{
		DrivenProcess.SetUpVars("E:\\Automation Framework\\I.A Tester - Framework\\I.A Tester - Workspace");
	}
	@Test
	public void testParse() 
	{	
		if (DrivenProcess.Connection("ALEJANDRO-BOTER", "OpenSourceFramework", "sa", "sql2008"))
		{
			do
			{
				DrivenProcess.OpenExecutions();
				try
				{
					if (DrivenProcess.RsExecutions.next())
					{
						DrivenProcess.InterfaceId = DrivenProcess.RsExecutions.getInt("InterfaceId");
						DrivenProcess.InterfaceName=DrivenProcess.RsExecutions.getString("InterfaceSciptName");
						DrivenProcess.UpdateExecution(DrivenProcess.InterfaceId);
						DrivenProcess.OpenResulsets(DrivenProcess.InterfaceId,DrivenProcess.InterfaceName);
						DrivenProcess.run();
						DrivenProcess.GenerateReport(DrivenProcess.InterfaceId,DrivenProcess.InterfaceName);
						DrivenProcess.CloseResulsets();
					}
				} catch (SQLException ex) {}
			}while (true);
		}
	}
}
