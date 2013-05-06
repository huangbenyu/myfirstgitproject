using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using System.Data;
using System.Data.OleDb;


namespace Exceltotxt
{
	class DBExcelHelper
	{

		private string _fileName = string.Empty;


        private OleDbConnection conn = null;  
		public UTF8Encoding utf8 = new UTF8Encoding();
		public string outstring;


		public DBExcelHelper(string filename)
		{
			_fileName = filename;
			outstring = "";


		}

		private bool IsNumeric(string number)
		{
			try
			{
				int.Parse(number);
				return true;
			}
			catch
			{
				return false;
			}
		}

		private bool IsFloat(string number)
		{
			try
			{
				float.Parse(number);
				return true;
			}
			catch
			{
				return false;
			}
		}
		public void ClearData()
		{
			outstring = "";
		}
		public string GetUtf8String(string ansistring)
		{

			byte[] data3 = utf8.GetBytes(ansistring);

			return utf8.GetString(data3, 0, data3.Length);
		}

		public bool LoadExcelFile()
		{
			string connStr = "";
			string fileType = System.IO.Path.GetExtension(_fileName);

			try
			{		
				if (string.IsNullOrEmpty(fileType)) 
				return false;

				if (fileType == ".xls")
					connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + _fileName + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
				else
					connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + _fileName + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";

				

				// 初始化连接，并打开  
				conn = new OleDbConnection(connStr);
				conn.Open();
                return true;

			}
			catch (System.Exception ex)
			{
				Program.Logger.Error("DBExcelHelper , Excel file  not exist", ex);
			}
 
			return false;
		}
		public int GetContent(string sheetName)
		{
			

			string sql_F =String.Format("Select * FROM [{0}$]", sheetName);


            OleDbDataAdapter da = new OleDbDataAdapter(sql_F, conn);  


			DataSet dsItem = new DataSet();
			da.Fill(dsItem, sheetName);

			//int m_maxcol = range.Columns.Count;
			//int m_row = range.Rows.Count;

			//if (m_row < 3)
			//{
			//    Program.Logger.ErrorFormat("sheet data Error, Name:{0},SheetName:{1}", _fileName, sheetName);
			//    return -1;
			//}
			foreach (DataTable table in dsItem.Tables)
			{
				int m_maxcol = table.Columns.Count;
				int m_row = table.Rows.Count;

				if (m_row < 3)
				{
					Program.Logger.ErrorFormat("sheet data Error, Name:{0},SheetName:{1}", _fileName, sheetName);
					return -1;
				}

				ArrayList fieldTypes = new ArrayList();

				for (int i = 0; i < m_maxcol; ++i)
				{
					String strtype = table.Rows[i][0].ToString();
					strtype = strtype.Trim();
					strtype = strtype.ToUpper();
					if (strtype == "INT" || strtype == "FLOAT" || strtype == "STRING")
					{
						fieldTypes.Add(strtype);
					}
					else
					{
						Program.Logger.ErrorFormat("Type, Name:{0},SheetName:{1} rol:{2}  Type Error ", _fileName, sheetName, i);
						return -1;
					}

				}
				//检测数据类型
				String tempstring;
				for (int i = 3; i < m_row; ++i)
				{
					for (int j = 0; j < m_maxcol; ++j)
					{
						tempstring = table.Rows[i][j].ToString();


						//数据
						if (fieldTypes[j].ToString() == "INT")
						{

							if (tempstring.Length != 0 && false == IsNumeric(tempstring))
							{
								Program.Logger.ErrorFormat("GetContent, Name:{0},SheetName:{1}  row:{2} rol:{3}  Type  [INT ] Error ", _fileName, sheetName, i, j);
								return -1;
							}
						}
						else if (fieldTypes[j].ToString() == "Float")
						{
							if (tempstring.Length != 0 && false == IsFloat(tempstring))
							{
								Program.Logger.ErrorFormat("GetContent, Name:{0},SheetName:{1}  row:{2}  rol:{3}  Type[Float] Error ", _fileName, sheetName, i, j);
								return -1;
							}
						}

					}
				}


				for (int i = 0; i < 3; ++i)
				{
					for (int j = 0; j < m_maxcol; ++j)
					{
						tempstring = "";
						tempstring = table.Rows[i][j].ToString();

						outstring += GetUtf8String(tempstring);



						if (j != m_maxcol)
						{
							outstring += "\t";
							outstring += GetUtf8String(tempstring);
						}

					}
					outstring += "\r\n";

				}

				for (int i = 3; i <m_row; ++i)
				{
					for (int j = 0; j <= m_maxcol; ++j)
					{
						tempstring = "";
						tempstring = table.Rows[i][j].ToString();


						//数据
						if (fieldTypes[j].ToString() == "INT" && tempstring.Length == 0)
						{
							outstring += "0";

						}
						else if (tempstring.Length == 0 && fieldTypes[j].ToString() == "Float")
						{
							outstring += "0";
						}
						else
						{
							outstring += GetUtf8String(tempstring);
						}

						if (j != m_maxcol)
						{
							outstring += "\t";

						}

					}
					outstring += "\r\n";

				}
			}

			


			return 0;
		}

		public void Close()
		{
			if (conn.State == ConnectionState.Open)
			{
				conn.Close();
				conn.Dispose();
			}  
		}

	}
}
