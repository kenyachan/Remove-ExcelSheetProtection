using System;
using System.IO;
using System.IO.Compression;
using System.Xml;

namespace Remove_ExcelSheetProtection
{
    class Program
    {
        static void Main(string[] args)
        {
            string file = args[0];
            string fileName = Path.GetFileNameWithoutExtension(file);
            string fileExtention = Path.GetExtension(file);
            string fileDirectory = Path.GetDirectoryName(file);
            string zipFile = $"{fileDirectory}\\{fileName}_{DateTime.Now.ToString("dd-MMM-yyy hh-mm-ss")}.zip";

            string workbookFolder = $"{fileDirectory}\\{Path.GetFileNameWithoutExtension(zipFile)}";
            string worksheetsDirectory;
            string[] sheets;

            // Convert .xlsx to .zip and extract
            File.Copy(file, zipFile);
            ZipFile.ExtractToDirectory($"{zipFile}",workbookFolder);
            File.Delete(zipFile);

            worksheetsDirectory= $"{workbookFolder}\\xl\\worksheets\\";
            sheets = Directory.GetFiles(worksheetsDirectory);

            // Remove sheet protection
            for(int i = 0; i < sheets.Length; i++) { RemoveProtection(sheets[i]); }

            ZipFile.CreateFromDirectory(workbookFolder, $"{fileDirectory}\\{fileName}_noProtection_{fileExtention}");
        }

        static void RemoveProtection(string worksheetPath)
        {
            XmlDocument document = new XmlDocument();

            document.Load(worksheetPath);
            
            XmlNamespaceManager xmlNamespaceManager = new XmlNamespaceManager(document.NameTable);
            xmlNamespaceManager.AddNamespace("worksheets", "http://schemas.openxmlformats.org/spreadsheetml/2006/main");

            XmlNode root = document.DocumentElement;
            XmlNodeList nodes = root.SelectNodes("//worksheets:sheetProtection",xmlNamespaceManager);
        
            foreach(XmlNode node in nodes) { node.ParentNode.RemoveChild(node); }

            document.Save(worksheetPath);
        }
    }
}
