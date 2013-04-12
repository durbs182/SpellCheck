using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Data;
using System.Runtime.Serialization;
using System.Drawing;
using System.Reflection;
using System.Threading;

using OfficeOpenXml.Style;
using OfficeOpenXml;

using SpellCheck.Common.XmlData;

namespace SpellCheck
{
    public class XmlToExcel
    {
    	static void Main(string[] args)
    	{
//    		string outFile = SaveXmlToExcel("test.xml",true);
//			SaveToXlsx("");
//    		Console.WriteLine(outFile);

			for(int i=1;i<120;i++)
			{
				Console.WriteLine(GetColumnLetter(i));
			}
			
    		Console.ReadLine();
    	}
    	
    	/// <summary>
    	/// Column enum for spreadsheet columns
    	/// </summary>
//    	private enum Columns
//    	{ 
//		     FileHandlerName = 	1,	
//			 FileHandlerDescription, 
//			 FilePath, 
//			 ErrorTotal, 
//			 StringId, 
//			 Locale, 
//			 IncorrectString,
//			 SuggestedCorrection, 
//			 Alert, 
//			 Error 
//    	}
    	

  		static string c_headerStyleName = "header";
		static string c_evenlineStyleName = "evenline";
		static string c_oddlineStyleName = "oddline";
		static object lockObject = new object();
  
    	
    	/// <summary>
    	/// Get the string format for the column from the given enum
    	/// </summary>
    	/// <param name="value"></param>
    	/// <returns></returns>
    	private static string GetColumnStringForamt(SpellingErrorData.Columns value)
		{
		    int pos = (int)Enum.Parse(typeof(SpellingErrorData.Columns), value.ToString());
    		
		    string colLetter = GetColumnLetter(pos);
		    
		    return colLetter + "{0}";
		}
    	
    	/// <summary>
    	/// Get the letter for the excel column
    	/// </summary>
    	/// <param name="pos"></param>
    	/// <returns></returns>
    	private static string GetColumnLetter(int pos)
    	{
		    char[] columnLetters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
    	    
    	    if(pos < 26)
    	    {
	    	    char col = columnLetters[pos-1];
	    	    return new string(col,1);    	    
    	    }
    	    else
    	    {    	   
    	    	double dbl = pos / 25;
    	    	int num = (int)Math.Round(dbl,MidpointRounding.AwayFromZero);
    	    	char col = columnLetters[num-1];

    	    	int pos2 = pos % 25;
    	    	char col2 = columnLetters[pos2];
    	    	
    	    	char[] colLetterArray = new char[]{col,col2};
    	    	
    	    	return new string(colLetterArray);
    	    }
    	}
    	
    	/// <summary>
    	/// Create style list for rows
    	/// </summary>
    	/// <param name="pck"></param>
    	private static void CreateStyles(ExcelPackage pck)
    	{
		    var headerStyle = pck.Workbook.Styles.CreateNamedStyle(c_headerStyleName);
			headerStyle.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
			headerStyle.Style.Font.Bold = true;
			headerStyle.Style.Font.Color.SetColor(Color.White);
			headerStyle.Style.Fill.BackgroundColor.SetColor(Color.CadetBlue);	
			
			var evenlineStyle = pck.Workbook.Styles.CreateNamedStyle(c_evenlineStyleName);
			evenlineStyle.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
			evenlineStyle.Style.Fill.BackgroundColor.SetColor(Color.Beige);

			var oddlineStyle = pck.Workbook.Styles.CreateNamedStyle(c_oddlineStyleName);
 			oddlineStyle.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
			oddlineStyle.Style.Fill.BackgroundColor.SetColor(Color.BurlyWood);
    	}
    	
    	/// <summary>
    	/// Add Workbook properties
    	/// </summary>
    	/// <param name="pck"></param>
    	private static void AddWorkbookProperties(ExcelPackage pck)
    	{
    		pck.Workbook.Properties.Author = "AutoSpellcheck.exe";
    		pck.Workbook.Properties.Title = "Spelling errors report";
    		pck.Workbook.Properties.Company = "Citrix";
    		pck.Workbook.Properties.Comments = "Auto generated spreadsheet from AutoSpellcheck.exe"; 
    	}
    	
    	static ExcelWorksheet c_ws;
    	static int c_rowCount = 1; // start with header row
    	static string c_cellRangeFormat;
    	static AutoResetEvent autoEvent;
    	
    	/// <summary>
    	/// Save the spellingerrors object to xlsx
    	/// </summary>
    	/// <param name="errors"></param>
    	/// <param name="resultPath"></param>
    	/// <returns></returns>
    	public static string SaveToXlsx(SpellingErrors errors, string resultPath)
    	{
            string excelFilePath = string.Format(resultPath,"xlsx");
            double colWidth = 25.0;
            double rowHeight = 15.0; 			
           
    		using (ExcelPackage pck = new ExcelPackage())
    		{       			
    			CreateStyles(pck);
    			AddWorkbookProperties(pck);
    			
    			c_ws = pck.Workbook.Worksheets.Add("SpellingErrors");   			
    			
    			c_ws.View.FreezePanes(2,1);
 
    			/// <summary>
    			/// have to set width and height to work around a bug in EPPLUS
    			/// if defaultRowHeight isn't set xml is maformed so Excel fails to load
    			/// </summary>    	
    			c_ws.defaultColWidth = colWidth;
    			c_ws.defaultRowHeight = rowHeight;
    			
    			// set range for columns we care about
    			c_cellRangeFormat = string.Format("{0}:{1}",GetColumnStringForamt(SpellingErrorData.Columns.FileHandlerName), GetColumnStringForamt(SpellingErrorData.Columns.Error));
    		
    			// use range to setup auto filter for each column
    			c_ws.Cells[string.Format(c_cellRangeFormat,c_rowCount)].AutoFilter = true;
    			
     			//set the style for the column headers 
     			//have to set each cell individually as cellRange.StyleName = headerStyleName;
    			// only sets the first cell in the range!!    			
   				foreach(SpellingErrorData.Columns col in Enum.GetValues(typeof(SpellingErrorData.Columns)))
    			{
    				string format = GetColumnStringForamt(col);
    				c_ws.Cells[string.Format(format,c_rowCount)].StyleName = c_headerStyleName;
    			}
   				
    		
   				//create header for column   				
   				foreach(SpellingErrorData.Columns column in SpellingErrorData.Columns.GetValues(typeof(SpellingErrorData.Columns)))
   				{
   				      AddCell(GetColumnStringForamt(column), SpellingErrorData.GetStringValue(column ));
   				}
           
	            c_rowCount++;
	            
	            //set up venathandlers for data received and 
	            SpellingErrorData.DataReceivedEvent += new SpellingErrorData.DataReceivedHandler(SpellingErrorData_DataReceivedEvent);
	            SpellingErrorData.DataCompleteEvent += new EventHandler(SpellingErrorData_DataCompleteEvent);
	            autoEvent = new AutoResetEvent(false);
	            
	            // start get data in a seperate thread to make sure the file is processed before we save the spreadsheet
	            Thread worker = new Thread(GetData);
	            worker.Start(errors);
	            
	            // wait for data to be received and datacomplete event fired
	            autoEvent.WaitOne();
	            
    			pck.SaveAs(new FileInfo(excelFilePath));
    		}
    		
    		return excelFilePath;
     	}
    	
    	/// <summary>
    	/// Delegate to handle starting data processing
    	/// </summary>
    	/// <param name="errors"></param>
    	static void GetData(object errors)
    	{
    		SpellingErrorData.GetData(errors as SpellingErrors);
    	}

    	/// <summary>
    	/// Delegate to handle datacomplete event 
    	/// </summary>
    	/// <param name="sender"></param>
    	/// <param name="e"></param>
    	static void SpellingErrorData_DataCompleteEvent(object sender, EventArgs e)
    	{
    		// unblock waiting thread
    		autoEvent.Set();
    	}

    	/// <summary>
    	/// Delegate to receive data 
    	/// writes each row to the spreadhseet
    	/// </summary>
    	/// <param name="sender"></param>
    	/// <param name="args"></param>
    	static void SpellingErrorData_DataReceivedEvent(object sender, SpellingErrorData.DataReceivedEventArgs args)
    	{
    		SpellingErrorData data = args.SpellingErrorData;
    		
    		// add a row for each spelling error
	        // alternate style for each row
			if(c_rowCount % 2 == 0)
			{
				c_ws.Cells[string.Format(c_cellRangeFormat,c_rowCount)].StyleName = c_evenlineStyleName;
			}
			else
			{
				c_ws.Cells[string.Format(c_cellRangeFormat,c_rowCount)].StyleName = c_oddlineStyleName;
			}	
		
		   	foreach(SpellingErrorData.Columns column in SpellingErrorData.Columns.GetValues(typeof(SpellingErrorData.Columns)))
			{
				object value = data.Data[column];				
			
				if(value != null && value.GetType().Equals(typeof(System.Int32)))
				{
					AddCell(GetColumnStringForamt(column), value, true);
				}
				else
				{
	   				AddCell(GetColumnStringForamt(column), value);
				}
		
			}
		   	
		   	c_rowCount ++;
	            
    	}
    	
		/// <summary>
		/// Helper method to add a value to a cell on the given worksheet
    	/// sets numeric format if isNum = true
		/// </summary>
		/// <param name="ws"></param>
		/// <param name="cellRangeFormat">A{0} etc</param>
		/// <param name="rowNum">Worsheet row number</param>
		/// <param name="cellValue">object containing cell value</param>
		/// <param name="isNum">Sets numeric format for cell if True</param>
    	private static void AddCell(string cellRangeFormat, object cellValue, bool isNum)
    	{  
    		if(isNum)
    		{
    			c_ws.Cells[string.Format(cellRangeFormat,c_rowCount)].Style.Numberformat.Format = "#,##0";
    			c_ws.Cells[string.Format(cellRangeFormat,c_rowCount)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
    		}

			if(cellValue != null)
			{
	     		c_ws.Cells[string.Format(cellRangeFormat,c_rowCount)].Value = cellValue.ToString();
			}
			else
			{
				c_ws.Cells[string.Format(cellRangeFormat,c_rowCount)].Value = "-";
			}
   		
    	} 
    	
    	private static void AddCell(string cellRangeFormat, object cellValue)
    	{    	
    		AddCell(cellRangeFormat,cellValue,false);
    	} 
    }
}
