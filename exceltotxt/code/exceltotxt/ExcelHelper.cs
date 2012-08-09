using System;
using System.Collections;
using System.IO;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

public class ExcelHelper
{
    private Excel._Application excelApp;
    private string fileName = string.Empty;
    private Excel.WorkbookClass wbclass;
    public ExcelHelper(string _filename)
    {
        excelApp = new Excel.Application();
        object objOpt = System.Reflection.Missing.Value;
        wbclass = (Excel.WorkbookClass)excelApp.Workbooks.Open(_filename, objOpt, false, objOpt, objOpt, objOpt, true, objOpt, objOpt, true, objOpt, objOpt, objOpt, objOpt, objOpt);
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
    /// <summary>
    /// 
    /// </summary>
    /// <param name="sheetName">sheet名称</param>
    /// <returns></returns>
    public int GetContent(string sheetName)
    {
        Excel.Worksheet sheet = GetWorksheetByName(sheetName);
        //获取A1 到AM24范围的单元格
        Excel.Range rang = sheet.get_Range("A1", "A3");

        int m_maxcol = 0;
        int m_row     =1;
        String tempstring;
        do{
            m_maxcol = m_maxcol + 1;
            rang = sheet.get_Range(sheet.Cells[m_row, m_maxcol], sheet.Cells[m_row, m_maxcol]);

            //tempstring = rang.Value2.ToString();
        }

        while (rang.Value2 != null);

        //读入类型数据
         
        String fieldTypes[512];  
        String strType="";

        for(int i =  0; i< m_maxcol ;++i )
        {

        }

        //String  value = sheet.Cells
        //读一个单元格内容
        //sheet.get_Range("A1", Type.Missing);
        //不为空的区域,列,行数目
        //   int l = sheet.UsedRange.Columns.Count;
        // int w = sheet.UsedRange.Rows.Count;
        //  object[,] dell = sheet.UsedRange.get_Value(Missing.Value) as object[,];
        System.Array values = (Array)rang.Cells.Value2;
        return 0;
    }

    public void Close()
    {
        excelApp.Quit();
        excelApp = null;
    }

}