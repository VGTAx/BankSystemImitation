using OfficeOpenXml;
using OfficeOpenXml.Style;
using InitHelperInformatMessage;
using System.Drawing;

namespace BankSystem
{
    internal class ExcelMethodGroup
    {
        public static bool CheckInfoXLSX(int value, string workbook, string worksheets)
        {
            try
            {
                byte[] bin = File.ReadAllBytes(workbook);
                MemoryStream memoryStream = new MemoryStream(bin);                
                ExcelPackage excelPackage = new ExcelPackage(memoryStream);
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[worksheets];

                for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
                {
                    for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.Start.Column; j++)
                    {
                        int tempValue = int.Parse(worksheet.Cells[i, j].Value.ToString());
                        if (tempValue == value)
                            return false;
                    }
                }
                return true;
            }
            catch (Exception)
            {
                return true;
            }
        }
        public static bool CheckAccNameAvailableXLSX(string login)
        {
            try
            {                
                byte[] bin = File.ReadAllBytes("ClientInfo.xlsx");
                MemoryStream memoryStream = new MemoryStream(bin);
                ExcelPackage package = new ExcelPackage(memoryStream);
                ExcelWorksheet worksheet = package.Workbook.Worksheets["ClientAccountInfo"];

                for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
                {
                    for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.Start.Column; j++)
                    {
                        string? tempLogin = worksheet.Cells[i, j].Value.ToString();
                        if (login == tempLogin)
                        {
                            MessageInformant.ErrorOutput($"Login \"{login}\" not available");
                            return false;
                        }
                    }
                }               
                return true;
            }
            catch (Exception)
            {
                return true;
            }
        }
        public static void EWorksheetAccountXLSX(IAccount account)
        {           
            ExcelPackage excelPackage = new ExcelPackage();
            excelPackage.Workbook.Properties.Author = "VGTAx";
            excelPackage.Workbook.Properties.Company = "PVG";
            excelPackage.Workbook.Properties.Title = "Title";
            excelPackage.Workbook.Properties.Created = DateTime.Now;

            FileInfo fileInfo = new FileInfo("ClientInfo.xlsx");
            if (fileInfo.Exists)
            {
                excelPackage = new ExcelPackage(fileInfo);
            }

            ExcelWorksheet? ClientInfoWS = excelPackage.Workbook.Worksheets["ClientInfo"];
            ExcelWorksheet? ClientAccInfoWS = excelPackage.Workbook.Worksheets["ClientAccountInfo"];

            if (ClientInfoWS == null && ClientAccInfoWS == null)
            {
                ClientInfoWS = excelPackage.Workbook.Worksheets.Add("ClientInfo");
                var ID = ClientInfoWS.Cells["A1:E1"];
                var name = ClientInfoWS.Cells[1,2];
                var surname = ClientInfoWS.Cells[1, 3];
                var age = ClientInfoWS.Cells[1, 4];
                var AmountOfMoney = ClientInfoWS.Cells[1, 5];

                ClientAccInfoWS = excelPackage.Workbook.Worksheets.Add("ClientAccountInfo");
                var login = ClientAccInfoWS.Cells["B1:C1"];
                var password = ClientAccInfoWS.Cells[1, 3];

                ID.IsRichText = true;
                ID.Style.WrapText = true;

                ID.Style.Border.Bottom.Style = ID.Style.Border.Right.Style =
                    ID.Style.Border.Left.Style = ID.Style.Border.Top.Style = ExcelBorderStyle.Medium;                

                var titleID = ID.RichText.Add("ID");
                titleID.Bold = true;
                titleID.FontName = "Calibri";
                titleID.Size = 14;

                var titleName = name.RichText.Add("Name");
                var titleSurname = surname.RichText.Add("Surname");
                var titleAge = age.RichText.Add("Age");
                var titleAmountOfMoney = AmountOfMoney.RichText.Add("Amount of Money");
                var titleLogin = login.RichText.Add("Login");
                var titlePassword = password.RichText.Add("Password");

                List<ExcelRichText> list = new List<ExcelRichText>();              
                list.Add(titleName);
                list.Add(titleAge);
                list.Add(titleSurname);
                list.Add(titleAmountOfMoney);
                list.Add(titleLogin);
                list.Add(titlePassword);

                foreach (var item in list)
                {
                    item.Bold = titleID.Bold;
                    item.FontName = titleID.FontName;
                    item.Size = titleID.Size;
                }
            }

            int rowClientWS = ClientInfoWS.Dimension.End.Row + 1;
            int colClientWS = ClientInfoWS.Columns.EndColumn;
           
            int rowAccountWS = ClientAccInfoWS.Dimension.End.Row+1;
            int colAccountWS = ClientAccInfoWS.Columns.EndColumn;
            

            if (rowClientWS > 2)
            {
                CheckInfoXLSX(account.ID,"ClientInfo.xlsx","ClientInfo");
                while (CheckInfoXLSX(account.ID, "ClientInfo.xlsx", "ClientInfo") == false)
                {
                    account.ID = new Random().Next(1, 9999);
                }
            }

            ClientAccInfoWS.Cells[ClientAccInfoWS.Dimension.Address].AutoFitColumns(15,20);          
            ClientInfoWS.Cells["A1:A10000"].AutoFitColumns(4);
            ClientInfoWS.Cells["B1:B10000"].AutoFitColumns(20);
            ClientInfoWS.Cells["C1:C10000"].AutoFitColumns(20);
            ClientInfoWS.Cells["D1:D10000"].AutoFitColumns(6);
            ClientInfoWS.Cells["E1:E10000"].AutoFitColumns(10);



            ClientInfoWS.Cells[rowClientWS, 1].Value = account.ID;
            ClientInfoWS.Cells[rowClientWS, 2].Value = account.personObj.Name;
            ClientInfoWS.Cells[rowClientWS, 3].Value = account.personObj.SurName;
            ClientInfoWS.Cells[rowClientWS, 4].Value = account.personObj.Age;
            ClientInfoWS.Cells[rowClientWS, 5].Value = account.AmountOfMoney;
            ClientAccInfoWS.Cells[rowAccountWS, 2].Value = account.Login;
            ClientAccInfoWS.Cells[rowAccountWS, 3].Value = account.Password;

            excelPackage.SaveAs("ClientInfo.xlsx");
        }
        public static double WithdrawMoneyXLSX(IAccount account, double amountMoney, double tempAmountMoney)
        {
            int tempRow = 0;
            byte[] bin = File.ReadAllBytes("ClientInfo.xlsx");
            MemoryStream memoryStream = new MemoryStream(bin);
            ExcelPackage package1 = new ExcelPackage(memoryStream);
            ExcelWorksheet worksheet = package1.Workbook.Worksheets["ClientInfo"];

            for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
            {
                for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.Start.Column; j++)
                {
                    int tempID = int.Parse(worksheet.Cells[i, j].Value.ToString());
                    if (account.ID == tempID)
                    {
                        tempRow = i;
                    }
                }
            }
            worksheet.Cells[tempRow, 5].Value = (amountMoney -= tempAmountMoney);
            package1.SaveAs("ClientInfo.xlsx");
            return amountMoney;
        }        
        public static double WithdrawMoneyAtmXLSX(IBankomat bankomat, double amountMoney, double tempAmountMoney)
        {
            int tempRow = 0;
            byte[] bin = File.ReadAllBytes("ATMInfo.xlsx");
            MemoryStream memoryStream = new MemoryStream(bin);
            ExcelPackage package1 = new ExcelPackage(memoryStream);
            ExcelWorksheet worksheet = package1.Workbook.Worksheets["ATM Info"];

            for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
            {
                for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.Start.Column; j++)
                {
                    int tempNumberATM = int.Parse(worksheet.Cells[i, j].Value.ToString());
                    if (bankomat.NumberATM == tempNumberATM)
                    {
                        tempRow = i;
                    }
                }
            }

            worksheet.Cells[tempRow, 3].Value = (amountMoney -= tempAmountMoney);
            package1.SaveAs("ATMInfo.xlsx");            
            return amountMoney;
        }
        public static void ExcelWorksheetAtmXLSX(IBankomat bankomat)
        {
            ExcelPackage packageATM = new ExcelPackage();
            packageATM.Workbook.Properties.Author = "VGTAx";
            packageATM.Workbook.Properties.Company = "PVG";
            packageATM.Workbook.Properties.Title = "Information about ATM";
            packageATM.Workbook.Properties.Created = DateTime.Now;

            FileInfo fileInfo = new FileInfo("ATMInfo.xlsx");
            if (fileInfo.Exists)
            {
                packageATM = new ExcelPackage(fileInfo);
            }

            ExcelWorksheet? worksheetATM = packageATM.Workbook.Worksheets["ATM Info"];

            int rowWS = 0;
            int colWS = 0;

            if (worksheetATM != null)
            {
                rowWS = worksheetATM.Dimension.End.Row+1;
                colWS = worksheetATM.Columns.EndColumn;
            }            

            if (worksheetATM == null)
            {               
                worksheetATM = packageATM.Workbook.Worksheets.Add("ATM Info");
                var numberATM = worksheetATM.Cells["A1"];
                var adressATM = worksheetATM.Cells["B1"];
                var MoneyATM = worksheetATM.Cells["C1"];                

                numberATM.IsRichText = true;
                numberATM.Style.WrapText = true;                

                var titleNumberATM = numberATM.RichText.Add("№ ATM");
                titleNumberATM.Bold = true;
                titleNumberATM.FontName = "Calibri";
                titleNumberATM.Size = 16;

                var titleAdressATM = adressATM.RichText.Add("Adress ATM");
                var titleMoneyATM = MoneyATM.RichText.Add("Amount of money");

                List<ExcelRichText> excels = new List<ExcelRichText>()
                {
                    titleMoneyATM,
                    titleAdressATM,
                };

                foreach (var item in excels)
                {
                    item.Bold = titleNumberATM.Bold;
                    item.FontName = titleNumberATM.FontName;
                    item.Size = titleNumberATM.Size;
                }

                rowWS = worksheetATM.Dimension.End.Row + 1;
                colWS = worksheetATM.Columns.EndColumn;
            }
            
            if (rowWS > 2)
            {
                CheckInfoXLSX(bankomat.NumberATM, "ATMInfo.xlsx", "ATM Info");
                while (CheckInfoXLSX(bankomat.NumberATM, "ATMInfo.xlsx", "ATM Info") == false)
                {
                    bankomat.NumberATM = new Random().Next(1, 9999);
                }
            }

            worksheetATM.Cells[1, 1, 1, colWS].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            worksheetATM.Cells[1, 1, 1, colWS].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            worksheetATM.Cells[1, 1, 1, colWS].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            worksheetATM.Cells[1, 1, rowWS, colWS].Style.Font.Size = 12;

            worksheetATM.Cells[1, 1, rowWS, 1].AutoFitColumns(8);
            worksheetATM.Cells[1, 2, rowWS, 2].AutoFitColumns(32);
            worksheetATM.Cells[1, 3, rowWS, 3].AutoFitColumns(15);            

            worksheetATM.Columns[1, colWS].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheetATM.Columns[1, colWS].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheetATM.Columns[1, colWS].Style.WrapText = true;

            worksheetATM.Cells[2, 1, rowWS, colWS].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            worksheetATM.Cells[2, 1, rowWS, colWS].Style.Border.Right.Style = ExcelBorderStyle.Medium;                                
            worksheetATM.Cells[2, 1, rowWS, colWS].Style.Border.Left.Style = ExcelBorderStyle.Medium;           

            worksheetATM.Cells[2, 3, rowWS, 3].Style.Font.Bold = true;
            worksheetATM.Cells[2, 3, rowWS, 3].Style.Font.Color.SetColor(Color.Red);

            worksheetATM.Cells[rowWS, 1].Value = bankomat.NumberATM;
            worksheetATM.Cells[rowWS, 2].Value = bankomat.Adress;
            worksheetATM.Cells[rowWS, 3].Value = bankomat.AmountOfMoneyATM;

            
            packageATM.SaveAs("ATMInfo.xlsx");
        }
        public static List<IAccount> LoadListAccXLSX()
        {
            List<IAccount> accountsList = new List<IAccount>();
            try
            {
                byte[]? bin = File.ReadAllBytes("ClientInfo.xlsx");
                MemoryStream memoryStream = new MemoryStream(bin);
                ExcelPackage excelPackage = new ExcelPackage(memoryStream);
                ExcelWorksheet clientInfo = excelPackage.Workbook.Worksheets["ClientInfo"];
                ExcelWorksheet accInfo = excelPackage.Workbook.Worksheets["ClientAccountInfo"];

                for (int i = clientInfo.Dimension.Start.Row + 1; i <= clientInfo.Dimension.End.Row/* + 1*/; i++)
                {
                    Account temp = new Account();
                    for (int j = clientInfo.Dimension.Start.Column; j < clientInfo.Dimension.End.Column + 1; j++)
                    {
                        if (j == 1)
                        {
                            temp.ID = int.Parse(clientInfo.Cells[i, j].Value.ToString());
                        }
                        if (j == 2)
                        {
                            temp.personObj.Name = clientInfo.Cells[i, j].Value.ToString();
                            temp.Login = accInfo.Cells[i, j].Value.ToString();
                        }
                        if (j == 3)
                        {
                            temp.personObj.SurName = clientInfo.Cells[i, j].Value.ToString();
                            temp.Password = accInfo.Cells[i, j].Value.ToString();
                        }
                        if (j == 4)
                            temp.personObj.Age = int.Parse(clientInfo.Cells[i, j].Value.ToString());
                        if (j == 5)
                        {
                            temp.AmountOfMoney = int.Parse(clientInfo.Cells[i, j]?.Value.ToString());
                        }

                    }
                    accountsList.Add(temp);
                }
                return accountsList;
            }
            catch (Exception)
            {
                return accountsList;
            }
        }
        public static List<IBankomat> LoadListAtmXLSX()
        {
            List<IBankomat> listATM = new List<IBankomat>();
            try
            {
                byte[] bin = File.ReadAllBytes("ATMInfo.xlsx");
                MemoryStream memoryStream = new MemoryStream(bin);
                ExcelPackage packageATM = new ExcelPackage(memoryStream);
                ExcelWorksheet worksheet = packageATM.Workbook.Worksheets["ATM Info"];
                for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
                {
                    Bankomat bankomat = new Bankomat();
                    for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                    {
                        if (j == 1)
                            bankomat.NumberATM = int.Parse(worksheet.Cells[i, j].Value.ToString());
                        if (j == 2)
                            bankomat.Adress = worksheet.Cells[i, j].Value.ToString();
                        if (j == 3)
                            bankomat.AmountOfMoneyATM = int.Parse(worksheet.Cells[i, j].Value.ToString());
                    }
                    listATM.Add(bankomat);
                }
                return listATM;
            }
            catch (Exception)
            {
                return listATM;
            }
        }
    }
}
