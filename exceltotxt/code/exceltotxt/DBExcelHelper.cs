using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace exceltotxt
{
	class DBExcelHelper
	{

		private string _fileName = string.Empty;


		OleDbConnection conn = null;  
		public UTF8Encoding utf8 = new UTF8Encoding();
		public string outstring;


		public DBExcelHelper(string filename)
		{
			_fileName = filename;
			outstring = "";


		}

		public bool LoadExcelFile()
		{
			string connStr = "";
			string fileType = System.IO.Path.GetExtension(_fileName);

			if (string.IsNullOrEmpty(fileType)) 
				return false;

			if (fileType == ".xls")
				connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + _fileName + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
			else
				connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + _fileName + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";

			string sql_F = "Select * FROM [{0}]";

			// 初始化连接，并打开  
			conn = new OleDbConnection(connStr);
			conn.Open();  


			return false;
		}

	}
}
