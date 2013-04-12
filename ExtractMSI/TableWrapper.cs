using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Globalization;

using SpellCheck.Common.Utility;
using Microsoft.Deployment.WindowsInstaller;

namespace ExtractMSI
{
	/// <summary>
	/// Represents a generic row in a table.
	/// </summary>
	public class TableRow
	{
		private readonly IDictionary _columns;

		private TableRow(IDictionary columns)
		{
			if (columns == null)
				throw new ArgumentNullException("columns");
			_columns = columns;
		}



        public static List<TableRow> GetRowsFromTable(Database msidb, string tableName)
		{
            if (!msidb.Tables.Contains(tableName) )
			{
				Lib.Log("Table name does {0} not exist Found.", tableName);
				return new List<TableRow>();
			}

			string query = string.Concat("SELECT * FROM `", tableName, "`");

            if (tableName == "Cabs")
            { 
            
            }

            var cols = msidb.Tables[tableName].Columns;
            var sql = msidb.Tables[tableName].SqlSelectString;
            
            List<TableRow> rows = new List<TableRow>();

            using (var view = msidb.OpenView(sql))
            {
                view.Execute();

                foreach (var values in view)
                {
                    Dictionary<string, object> valueCollection = new Dictionary<string, object>();
                    for (int cIndex = 0; cIndex < cols.Count; cIndex++)
                    {
                        string s;
                        object o;

                        try
                        {
                            s = cols[cIndex].Name;
                            try
                            {
                                o = values.GetString(s); ;
                            }
                            catch (Microsoft.Deployment.WindowsInstaller.InstallerException inex)
                            {
                                o = string.Empty;
                                Lib.Log(inex.Message);
                            }
                            //Lib.DebugLog("{0} {1}", s, o);
                            valueCollection.Add(s, o);
                        }
                        
                        catch (AccessViolationException ex)
                        {
                            Lib.Log(ex.StackTrace);
                        }
                        catch (NullReferenceException nrex)
                        {
                            Lib.Log(nrex.StackTrace);
                        }
                        catch (IndexOutOfRangeException ioorex)
                        {
                            Lib.Log(ioorex.StackTrace);
                        }
                    }

                    try
                    {
                        rows.Add(new TableRow(valueCollection));
                    }
                    catch (IndexOutOfRangeException ioorex)
                    {
                        Lib.Log(ioorex.Message);
                    }
                }

                view.Close();
            }
			return rows;
		}

		public string GetString(string columnName)
		{
			return Convert.ToString(GetValue(columnName));
		}
		public object GetValue(string columnName)
		{
            object obj = _columns[columnName];
            return obj;
		}

		public Int32 GetInt32(string columnName)
		{
			return Convert.ToInt32(GetValue(columnName));
		}
	}
}