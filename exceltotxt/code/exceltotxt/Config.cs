
using System;
using System.Collections;
using System.IO;
using System.Data;

using System.Collections.Generic;
using System.Text;

using System.Xml;

public class ExcelData
{

    public string fileName;
    public string sheetName;
    public bool  client;
    public string GetFileName()
    {
        return fileName;
    }

    public string GetSheetName()
    {
        return fileName;
    }

    public void Parse(XmlNode node)
    {
        XmlAttribute xmlName = node.Attributes["FileName"];
        if (xmlName == null)
        {
            throw new Exception("VI_Define no [FileName] atrribute!");
        }
        fileName = xmlName.Value;

        XmlAttribute xmlType = node.Attributes["SheetName"];
        if (xmlType == null)
        {
            throw new Exception("VI_Define no [SheetName] atrribute!");
        }
        sheetName = xmlType.Value;


        XmlAttribute xmlClient = node.Attributes["Client"];
        if (xmlClient == null)
        {
            throw new Exception("VI_Define no [Client] atrribute!");
        }
        client = Convert.ToBoolean(xmlClient.Value);

    }
}


public class ExcelConfigFile
{
    public System.Collections.ArrayList excelDatalist = new System.Collections.ArrayList();
    public string ServerPath;
    public string ClientPath;

    public void Parse(XmlNode doc)
    {

        XmlNodeList roots = doc.SelectNodes("ConfigData");
        if (roots.Count == 0)
        {
            throw new Exception("ExcelFile not found!");
        }
        if (roots.Count > 1)
        {
            throw new Exception("ExcelFile  Error!");
        }

        XmlNode root = roots[0];

        XmlAttribute xmlName = root.Attributes["Serverpath"];
        if (xmlName == null)
        {
            throw new Exception("protocol no [Serverpath] atrribute!");
        }

        ServerPath = xmlName.Value;


        xmlName = root.Attributes["ClientPath"];
        if (xmlName == null)
        {
            throw new Exception("protocol no [ClientPath] atrribute!");
        }

        ClientPath = xmlName.Value; ;

        //define

        XmlNodeList xmldefinesList = root.SelectNodes("ExcelFile");
        for (int i = 0; i < xmldefinesList.Count; i++)
        {
            XmlNode xmldefines = xmldefinesList[i];
            ExcelData excelData = new ExcelData();
            excelData.Parse(xmldefines);

            excelDatalist.Add(excelData);


        }
        
    }


}
