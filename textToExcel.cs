using System;
using System.IO;
using System.Reflection;
using Microsoft.Office.Interop.Excel;

public class rawParser
{
	static void Main()
	{
		try{
			using (StreamReader sr = new StreamReader("test.txt"))
			{
				string raw = sr.ReadToEnd();
				string[] scans = raw.Split(';');

				foreach(string s in scans)

				{

					string scan = s.Trim();
	   				Console.WriteLine("We have {0}", scan);
				}
			}
		}
		catch(Exception e)
		{
			Console.WriteLine("File error: ");
			Console.WriteLine(e.Message);
		}
		Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
		if(excel == null)
		{
			Console.WriteLine("Excel could not be started.");
			return;
		}
		excel.Visible = true;

		Workbook wb = excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
		Worksheet ws = (Worksheet)wb.Worksheets[1];

		if( ws == null )
		{
			Console.Writeline("Worksheet could not be created.");
		}

		Range aRange = ws.get_Range("C1", "C7");

		if(aRange == null)
		{
			Console.WriteLine("Could not get range.");
		}

		Object[] args = new Object[1];
        args[0] = 6;
        aRange.GetType().InvokeMember("Value", BindingFlags.SetProperty, null, aRange, args);
    
        aRange.Value2 = 8;
	}
}
