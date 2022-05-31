using OfficeOpenXml;
using OfficeOpenXml.Style;
using InitHelperInformatMessage;

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

            int rowClient = 1;
            int rowAccount = 1;

            for (int i = ClientInfoWS.Dimension.Start.Row; i <= ClientInfoWS.Dimension.End.Row; i++)
            {
                rowClient++;
            }
            for (int i = ClientAccInfoWS.Dimension.Start.Row; i <= ClientAccInfoWS.Dimension.End.Row; i++)
            {
                rowAccount++;
            }

            if (rowClient > 2)
            {
                CheckInfoXLSX(account.ID,"ClientInfo.xlsx","ClientInfo");
                while (CheckInfoXLSX(account.ID, "ClientInfo.xlsx", "ClientInfo") == false)
                {
                    account.ID = new Random().Next(1, 9999);
                }
            }

            ClientInfoWS.Cells[rowClient, 1].Value = account.ID;
            ClientInfoWS.Cells[rowClient, 2].Value = account.personObj.Name;
            ClientInfoWS.Cells[rowClient, 3].Value = account.personObj.SurName;
            ClientInfoWS.Cells[rowClient, 4].Value = account.personObj.Age;
            ClientInfoWS.Cells[rowClient, 5].Value = account.AmountOfMoney;
            ClientAccInfoWS.Cells[rowAccount, 2].Value = account.Login;
            ClientAccInfoWS.Cells[rowAccount, 3].Value = account.Password;

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
            if(fileInfo.Exists)
            {
                packageATM = new ExcelPackage(fileInfo);
            }

            ExcelWorksheet? worksheetATM = packageATM.Workbook.Worksheets["ATM Info"];

            if(worksheetATM == null)
            {
                worksheetATM = packageATM.Workbook.Worksheets.Add("ATM Info");
                var numberATM = worksheetATM.Cells["A1"];
                var adressATM = worksheetATM.Cells["B1"];
                var MoneyATM = worksheetATM.Cells["C1"];

                numberATM.IsRichText = true;
                numberATM.Style.WrapText = true;
                numberATM.Style.Border.Bottom.Style = numberATM.Style.Border.Top.Style =
                    numberATM.Style.Border.Right.Style = numberATM.Style.Border.Left.Style = ExcelBorderStyle.Medium;

                var titleNumberATM = numberATM.RichText.Add("№ ATM");
                titleNumberATM.Bold = true;
                titleNumberATM.FontName = "Calibri";
                titleNumberATM.Size = 14;

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
            }

            int rowClient = 1;

            for (int i = worksheetATM.Dimension.Start.Row; i <= worksheetATM.Dimension.End.Row; i++)
            {
                rowClient++;
            }
            if (rowClient > 2)
            {
                CheckInfoXLSX(bankomat.NumberATM, "ATMInfo.xlsx", "ATM Info");
                while (CheckInfoXLSX(bankomat.NumberATM, "ATMInfo.xlsx", "ATM Info") == false)
                {
                    bankomat.NumberATM = new Random().Next(1, 9999);
                }
            }

            worksheetATM.Cells[rowClient,1].Value = bankomat.NumberATM;
            worksheetATM.Cells[rowClient, 2].Value = bankomat.Adress;
            worksheetATM.Cells[rowClient, 3].Value = bankomat.AmountOfMoneyATM;

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
