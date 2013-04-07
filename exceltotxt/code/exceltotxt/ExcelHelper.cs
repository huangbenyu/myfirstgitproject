﻿using System;
using System.Collections;
using System.IO;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

using System.Collections.Generic;
using System.Text;



public class ExcelHelper
{
    private Excel._Application excelApp;
    private string _fileName = string.Empty;
    private Excel.WorkbookClass wbclass;
    public ExcelHelper(string filename)
    {
        _fileName = filename;
        excelApp = new Excel.Application();
        object objOpt = System.Reflection.Missing.Value;
        wbclass = (Excel.WorkbookClass)excelApp.Workbooks.Open(_fileName, objOpt, false, objOpt, objOpt, objOpt, true, objOpt, objOpt, true, objOpt, objOpt, objOpt, objOpt, objOpt);
    }
    /// <summary>
    /// 所有sheet的名称列表
    /// </summary>
    /// <returns></returns>
    //public List<string> GetSheetNames()
    //{
    //    List<string> list = new List<string>();
    //    Excel.Sheets sheets = wbclass.Worksheets;
    //    string sheetNams = string.Empty;
    //    foreach (Excel.Worksheet sheet in sheets)
    //    {
    //        list.Add(sheet.Name);
    //    }
    //    return list;
    //}
    public Excel.Worksheet GetWorksheetByName(string name)
    {
        Excel.Worksheet sheet = null;
        Excel.Sheets sheets = wbclass.Worksheets;
        foreach (Excel.Worksheet s in sheets)
        {
            if (s.Name == name)
            {
                sheet = s;
                break;
            }
        }
        return sheet;
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


    /// <summary>
    /// 
    /// </summary>
    /// <param name="sheetName">sheet名称</param>
    /// <returns></returns>
    public int GetContent(string sheetName)
    {
        Excel.Worksheet sheet = GetWorksheetByName(sheetName);
        //获取A1 到AM24范围的单元格
       // Excel.Range rang = sheet.get_Range("A1", "A3");

        Excel.Range   range = sheet.UsedRange;
        int m_maxcol    = range.Columns.Count;
        int m_row       = range.Rows.Count;

        if (m_row < 3 )
        {
            Console.WriteLine("GetContent, Name:{0},SheetName:{1}", _fileName, sheetName);
            return -1;
        }
        ArrayList fieldTypes = new ArrayList();

        for (int i = 1; i <= m_maxcol; ++i)
        {
            String strtype = Convert.ToString((range.Cells[1, i] as Excel.Range).Value2);
            strtype = strtype.Trim();
            strtype =strtype.ToUpper();
             if(strtype == "INT" || strtype == "FLOAT" || strtype == "STRING" )
            {
                fieldTypes.Add(strtype);
            }
            else
            {
                Console.WriteLine("GetContent, Name:{0},SheetName:{1} rol:{2}  Type Error ", _fileName, sheetName,i);
                return -1;
            }

        }
        //检测数据类型
         String tempstring;
         for (int i = 3; i <= m_row; ++i)
        {
            for (int j = 1; j <= m_maxcol; ++j )
            {
                tempstring = Convert.ToString((range.Cells[i,j] as Excel.Range).Value2);

             
                    //数据
                    if (fieldTypes[j-1].ToString() =="INT")
                    {
                        if (false == IsNumeric(tempstring))
                        {
                             Console.WriteLine("GetContent, Name:{0},SheetName:{1}  row:{2} rol:{3}  Type Error ", _fileName, sheetName,i,j);
                             return -1;
                        }
                    }
                    else if( fieldTypes[j-1].ToString() == "Float")
                    {
                         if (false == IsFloat(tempstring))
                        {
                             Console.WriteLine("GetContent, Name:{0},SheetName:{1}  row:{2}  rol:{3}  Type Error ", _fileName, sheetName,i,j);
                             return -1;
                        }
                    }
                    
             }
         }


        FileStream fs = new FileStream("H:\\test.txt", FileMode.Create);
        //获得字节数组
        

      
        for (int i = 1; i <= m_row; ++i)
        {
            for (int j = 1; j <= m_maxcol; ++j )
            {
                tempstring = Convert.ToString((range.Cells[i,j] as Excel.Range).Value2);

                
                
                 byte[] data = new UTF8Encoding().GetBytes(tempstring);
                    //开始写入
                    fs.Write(data, 0, data.Length);
                
  
                 if (j != m_maxcol)
                 {
                     tempstring = "\t";
                     byte[] data2 = new UTF8Encoding().GetBytes(tempstring);
                     //开始写入
                     fs.Write(data2, 0, data2.Length);
                 }
                 
            }


            tempstring = "\r\n";
            byte[] data3 = new UTF8Encoding().GetBytes(tempstring);
            //开始写入
            fs.Write(data3, 0, data3.Length);

        }

        //清空缓冲区、关闭流
        fs.Flush();
        fs.Close();

  
   
      
        return 0;
    }

    public void Close()
    {
        excelApp.Quit();
        excelApp = null;
    }

}