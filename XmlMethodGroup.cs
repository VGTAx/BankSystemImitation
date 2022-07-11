using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace BankSystem
{
    internal class XmlMethodGroup
    {
        public static bool CheckInfoXml(int value, string xmlFile, string worksheets, string WBPass = "",
           string WSPass = "")
        {
            //try
            //{
            //    XDocument xDoc = XDocument.Load(xmlFile);
            //    foreach (XElement xElement in xDoc.Elements()
            //    {

            //    }


            //    ///open file for read
            //    byte[] bin = File.ReadAllBytes(xmlFile);
            //    MemoryStream memoryStream = new MemoryStream(bin);
            //    ExcelPackage excelPackage = new ExcelPackage(memoryStream, GetPassword(WBPass, WSPass));
            //    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[worksheets];
            //    ///look for the same value
            //    for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
            //    {
            //        for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.Start.Column; j++)
            //        {
            //            int tempValue = int.Parse(worksheet.Cells[i, j].Value.ToString());
            //            if (tempValue == value)
            //                return false;
            //        }
            //    }
            //    return true;
            //}
            //catch (Exception)
            //{
            //    return true;
            //}
            return true;
        }

        public static void OpenOrCreateXmlAccountFile(IAccount account)
        {
            XDocument xAccountDocument = new XDocument();
            XDocument xAccountLoginInfoDocument = new XDocument();
            
            XElement xAccLoginInfoRootElem = new XElement("AccountInfo");
            XElement rootAccRootElem = new XElement("Accounts");

            XElement xAccPerson = new XElement("Person",
                                                new XAttribute("ID", account.ID),
                                                new XElement("Name", account.person.Name),
                                                new XElement("Surname", account.person.SurName),
                                                new XElement("Age", account.person.Age));

            XElement  xAccLoginInfo = new XElement("Account",
                                new XAttribute("ID", account.ID),
                                new XElement("Login", account.Login),
                                new XElement("Password", account.Password));               

            try
            {
                xAccountDocument = XDocument.Load("Accounts.xml");
                xAccountDocument.Element("Accounts").Add(xAccPerson);

                xAccountLoginInfoDocument = XDocument.Load("AccountLoginInfo.xml");                
                xAccountLoginInfoDocument.Element("AccountInfo").Add(xAccLoginInfo);
            }
            catch (Exception)
            {            
                xAccountDocument.Add(rootAccRootElem);                
                xAccountLoginInfoDocument.Add(xAccLoginInfoRootElem);
            }
            xAccountLoginInfoDocument.Save("AccountLoginInfo.xml");
            xAccountDocument.Save("Accounts.xml");                 
        }

    }
}
