 using System;
 using System.IO;


class rawParser
{
	public static void Main()
	{
		try{
			using (StreamReader sr = new StreamReader("test.txt"))
			{
				string raw = sr.ReadAllText();
				string[] scans = raw.Split(';');
				foreach(string s in scans)
				{
   					// Consume
	   				Console.WriteLine("We have {0}", s);
				}
			}
		}
		catch(Exception e)
		{
			Console.WriteLine("File error: ");
			Console.WriteLine(e.Message);
		}
	}
}