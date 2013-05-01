
using System;
using System.Collections;
using System.IO;
using System.Data;

using System.Collections.Generic;
using System.Text;

using System.Xml;

public class Sheetdata
{
    public string sheetName;
    public bool client;

    public void Parse(XmlNode node)
    {
        

        XmlAttribute xmlType = node.Attributes["Name"];
        if (xmlType == null)
        {
            throw new Exception("Sheetdata no Name] atrribute!");
        }
        sheetName = xmlType.Value;


        XmlAttribute xmlClient = node.Attributes["Client"];
        if (xmlClient == null)
        {
            throw new Exception("Sheetdata no [Client] atrribute!");
        }
        client = Convert.ToBoolean(xmlClient.Value);

    }

}
public class ExcelData
{
    public string fileName;

    public System.Collections.ArrayList sheetDatalist = new System.Collections.ArrayList();

    public string GetFileName()
    {
        return fileName;
    }

    public void Parse(XmlNode node)
    {
        XmlAttribute xmlName = node.Attributes["FileName"];
        if (xmlName == null)
        {
            throw new Exception("ExcelData no [FileName] atrribute!");
        }
        fileName = xmlName.Value;


        XmlNodeList xmldefinesList = node.SelectNodes("SheetName");
        for (int i = 0; i < xmldefinesList.Count; i++)
        {
            XmlNode xmldefines = xmldefinesList[i];
            Sheetdata sheetData = new Sheetdata();
            sheetData.Parse(xmldefines);

            sheetDatalist.Add(sheetData);

        }

       
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

        ClientPath = xmlName.Value;

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
