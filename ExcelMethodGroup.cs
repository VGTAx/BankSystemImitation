using InitHelperInformatMessage;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Rijndael256;
using System.Drawing;

namespace BankSystem
{
    public class ExcelMethodGroup
    {
        /// <summary>
        /// Сhecking for repetitions of the passed argument
        /// </summary>
        /// <param name="value">Сhecked value (ID or №ATM)</param>
        /// <param name="workbook">Name workbook with format(xlsx)</param>
        /// <param name="worksheets">Name worksheet in the workbook</param>
        /// <param name="WBPass">Name password workbook </param>
        /// <param name="WSPass">Name password worksheet</param>
        /// <returns></returns>
        public static bool CheckInfoXLSX(int value, string workbook, string worksheets, string WBPass = "",
            string WSPass = "")
        {
            try
            {
                ///open file for read
                byte[] bin = File.ReadAllBytes(workbook);
                MemoryStream memoryStream = new MemoryStream(bin);
                ExcelPackage excelPackage = new ExcelPackage(memoryStream, GetPassword(WBPass, WSPass));
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[worksheets];
                ///look for the same value
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

        /// <summary>
        /// Сhecking for availability Account name (Login)
        /// </summary>
        /// <param name="login">Checked value</param>
        /// <returns></returns>
        public static bool CheckAccNameAvailableXLSX(string login)
        {
            try
            {
                ///open file for read
                byte[] bin = File.ReadAllBytes("ClientInfo.xlsx");
                MemoryStream memoryStream = new MemoryStream(bin);
                ExcelPackage package = new ExcelPackage(memoryStream, GetPassword("PasswordClient.xlsx", "password"));
                ExcelWorksheet worksheet = package.Workbook.Worksheets["ClientAccountInfo"];
                ///look for the same login
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

        /// <summary>
        /// Create or open workbook with Account info
        /// </summary>
        /// <param name="account"></param>
        public static void WorksheetAccountXLSX(IAccount account)
        {
            ExcelPackage excelPackage = new ExcelPackage();
            excelPackage.Workbook.Properties.Author = "VGTAx";
            excelPackage.Workbook.Properties.Company = "PVG";
            excelPackage.Workbook.Properties.Title = "Title";
            excelPackage.Workbook.Properties.Created = DateTime.Now;

            FileInfo fileInfo = new FileInfo("ClientInfo.xlsx");
            if (fileInfo.Exists)
            {
                //Open excel file
                excelPackage = new ExcelPackage(fileInfo, GetPassword("PasswordClient.xlsx", "password"));
            }

            ExcelWorksheet? ClientInfoWS = excelPackage.Workbook.Worksheets["ClientInfo"];
            ExcelWorksheet? ClientAccInfoWS = excelPackage.Workbook.Worksheets["ClientAccountInfo"];

            if (ClientInfoWS == null && ClientAccInfoWS == null)
            {
                //create worksheet "Client Info"
                ClientInfoWS = excelPackage.Workbook.Worksheets.Add("ClientInfo");
                var ID = ClientInfoWS.Cells["A1:E1"];
                var name = ClientInfoWS.Cells[1, (int)EnumClient.Name];
                var surname = ClientInfoWS.Cells[1, (int)EnumClient.Surname];
                var age = ClientInfoWS.Cells[1, (int)EnumClient.Age];
                var AmountOfMoney = ClientInfoWS.Cells[1, (int)EnumClient.Money];
                //create worksheet "Account Info" (login & password)
                ClientAccInfoWS = excelPackage.Workbook.Worksheets.Add("ClientAccountInfo");
                var login = ClientAccInfoWS.Cells["B1:C1"];
                var password = ClientAccInfoWS.Cells[1, (int)EnumAcc.Password];

                ID.IsRichText = true;
                ID.Style.WrapText = true;
                //font style setting 
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

                List<ExcelRichText> list = new List<ExcelRichText>()
                {
                    titleName, titleSurname, titleAge, titleAmountOfMoney, titleLogin, titlePassword
                };

                foreach (var item in list)
                {
                    item.Bold = titleID.Bold;
                    item.FontName = titleID.FontName;
                    item.Size = titleID.Size;
                }
            }

            int rowClientWS = ClientInfoWS.Dimension.End.Row + 1;
            int colClientWS = ClientInfoWS.Columns.EndColumn;

            int rowAccountWS = ClientAccInfoWS.Dimension.End.Row + 1;

            //check ID account for uniqueness
            if (rowClientWS > 2)
            {
                while (CheckInfoXLSX(account.ID, "ClientInfo.xlsx", "ClientInfo", "PasswordClient.xlsx",
                    "password") == false)
                {
                    account.ID = new Random().Next(1, 9999);
                }
            }
            ///table(font and border) settings
            ClientInfoWS.Cells[1, 1, 1, colClientWS].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            ClientInfoWS.Cells[1, 1, 1, colClientWS].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            ClientInfoWS.Cells[1, 1, 1, colClientWS].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            ClientInfoWS.Cells[1, 1, rowClientWS, colClientWS].Style.Font.Size = 12;

            ClientInfoWS.Cells[1, (int)EnumClient.ID, rowClientWS, (int)EnumClient.ID].AutoFitColumns(6);
            ClientInfoWS.Cells[1, (int)EnumClient.Name, rowClientWS, (int)EnumClient.Name].AutoFitColumns(20);
            ClientInfoWS.Cells[1, (int)EnumClient.Surname, rowClientWS, (int)EnumClient.Surname].AutoFitColumns(20);
            ClientInfoWS.Cells[1, (int)EnumClient.Age, rowClientWS, (int)EnumClient.Age].AutoFitColumns(6);
            ClientInfoWS.Cells[1, (int)EnumClient.Money, rowClientWS, (int)EnumClient.Money].AutoFitColumns(10);

            ClientInfoWS.Columns[1, colClientWS].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ClientInfoWS.Columns[1, colClientWS].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ClientInfoWS.Columns[1, colClientWS].Style.WrapText = true;

            ClientInfoWS.Cells[2, 1, rowClientWS, colClientWS].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            ClientInfoWS.Cells[2, 1, rowClientWS, colClientWS].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            ClientInfoWS.Cells[2, 1, rowClientWS, colClientWS].Style.Border.Right.Style = ExcelBorderStyle.Medium;

            ClientInfoWS.Cells[2, (int)EnumClient.Money, rowClientWS, (int)EnumClient.Money].Style.Font.Color.SetColor(Color.Red);
            ClientInfoWS.Cells[2, (int)EnumClient.Money, rowClientWS, (int)EnumClient.Money].Style.Font.Bold = true;

            ClientInfoWS.Cells[rowClientWS, (int)EnumClient.ID].Value = account.ID;
            ClientInfoWS.Cells[rowClientWS, (int)EnumClient.Name].Value = account.person.Name;
            ClientInfoWS.Cells[rowClientWS, (int)EnumClient.Surname].Value = account.person.SurName;
            ClientInfoWS.Cells[rowClientWS, (int)EnumClient.Age].Value = account.person.Age;
            ClientInfoWS.Cells[rowClientWS, (int)EnumClient.Money].Value = account.AmountOfMoney;

            ClientAccInfoWS.Cells[rowAccountWS, (int)EnumAcc.Login].Value = account.Login;
            ClientAccInfoWS.Cells[rowAccountWS, (int)EnumAcc.Password].Value = account.Password;
            //save and close excel file
            excelPackage.SaveAs("ClientInfo.xlsx",
                SetPassword("PasswordClient.xlsx", "password"));
        }

        public static async void WorksheetAccountXLSXAsync(IAccount account)
        {
            ExcelPackage excelPackage = new ExcelPackage();
            excelPackage.Workbook.Properties.Author = "VGTAx";
            excelPackage.Workbook.Properties.Company = "PVG";
            excelPackage.Workbook.Properties.Title = "Title";
            excelPackage.Workbook.Properties.Created = DateTime.Now;

            FileInfo fileInfo = new FileInfo("ClientInfo.xlsx");
            if (fileInfo.Exists)
            {
                //Open excel file
                excelPackage = new ExcelPackage(fileInfo, GetPassword("PasswordClient.xlsx", "password"));
            }

            ExcelWorksheet? ClientInfoWS = excelPackage.Workbook.Worksheets["ClientInfo"];
            ExcelWorksheet? ClientAccInfoWS = excelPackage.Workbook.Worksheets["ClientAccountInfo"];

            if (ClientInfoWS == null && ClientAccInfoWS == null)
            {
                //create worksheet "Client Info"
                ClientInfoWS = excelPackage.Workbook.Worksheets.Add("ClientInfo");
                var ID = ClientInfoWS.Cells["A1:E1"];
                var name = ClientInfoWS.Cells[1, (int)EnumClient.Name];
                var surname = ClientInfoWS.Cells[1, (int)EnumClient.Surname];
                var age = ClientInfoWS.Cells[1, (int)EnumClient.Age];
                var AmountOfMoney = ClientInfoWS.Cells[1, (int)EnumClient.Money];
                //create worksheet "Account Info" (login & password)
                ClientAccInfoWS = excelPackage.Workbook.Worksheets.Add("ClientAccountInfo");
                var login = ClientAccInfoWS.Cells["B1:C1"];
                var password = ClientAccInfoWS.Cells[1, (int)EnumAcc.Password];

                ID.IsRichText = true;
                ID.Style.WrapText = true;
                //font style setting 
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

                List<ExcelRichText> list = new List<ExcelRichText>()
                {
                    titleName, titleSurname, titleAge, titleAmountOfMoney, titleLogin, titlePassword
                };

                foreach (var item in list)
                {
                    item.Bold = titleID.Bold;
                    item.FontName = titleID.FontName;
                    item.Size = titleID.Size;
                }
            }

            int rowClientWS = ClientInfoWS.Dimension.End.Row + 1;
            int colClientWS = ClientInfoWS.Columns.EndColumn;

            int rowAccountWS = ClientAccInfoWS.Dimension.End.Row + 1;

            //check ID account for uniqueness
            if (rowClientWS > 2)
            {
                while (CheckInfoXLSX(account.ID, "ClientInfo.xlsx", "ClientInfo", "PasswordClient.xlsx",
                    "password") == false)
                {
                    account.ID = new Random().Next(1, 9999);
                }
            }
            ///table(font and border) settings
            ClientInfoWS.Cells[1, 1, 1, colClientWS].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            ClientInfoWS.Cells[1, 1, 1, colClientWS].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            ClientInfoWS.Cells[1, 1, 1, colClientWS].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            ClientInfoWS.Cells[1, 1, rowClientWS, colClientWS].Style.Font.Size = 12;

            ClientInfoWS.Cells[1, (int)EnumClient.ID, rowClientWS, (int)EnumClient.ID].AutoFitColumns(6);
            ClientInfoWS.Cells[1, (int)EnumClient.Name, rowClientWS, (int)EnumClient.Name].AutoFitColumns(20);
            ClientInfoWS.Cells[1, (int)EnumClient.Surname, rowClientWS, (int)EnumClient.Surname].AutoFitColumns(20);
            ClientInfoWS.Cells[1, (int)EnumClient.Age, rowClientWS, (int)EnumClient.Age].AutoFitColumns(6);
            ClientInfoWS.Cells[1, (int)EnumClient.Money, rowClientWS, (int)EnumClient.Money].AutoFitColumns(10);

            ClientInfoWS.Columns[1, colClientWS].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            ClientInfoWS.Columns[1, colClientWS].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ClientInfoWS.Columns[1, colClientWS].Style.WrapText = true;

            ClientInfoWS.Cells[2, 1, rowClientWS, colClientWS].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            ClientInfoWS.Cells[2, 1, rowClientWS, colClientWS].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            ClientInfoWS.Cells[2, 1, rowClientWS, colClientWS].Style.Border.Right.Style = ExcelBorderStyle.Medium;

            ClientInfoWS.Cells[2, (int)EnumClient.Money, rowClientWS, (int)EnumClient.Money].Style.Font.Color.SetColor(Color.Red);
            ClientInfoWS.Cells[2, (int)EnumClient.Money, rowClientWS, (int)EnumClient.Money].Style.Font.Bold = true;

            ClientInfoWS.Cells[rowClientWS, (int)EnumClient.ID].Value = account.ID;
            ClientInfoWS.Cells[rowClientWS, (int)EnumClient.Name].Value = account.person.Name;
            ClientInfoWS.Cells[rowClientWS, (int)EnumClient.Surname].Value = account.person.SurName;
            ClientInfoWS.Cells[rowClientWS, (int)EnumClient.Age].Value = account.person.Age;
            ClientInfoWS.Cells[rowClientWS, (int)EnumClient.Money].Value = account.AmountOfMoney;

            ClientAccInfoWS.Cells[rowAccountWS, (int)EnumAcc.Login].Value = account.Login;
            ClientAccInfoWS.Cells[rowAccountWS, (int)EnumAcc.Password].Value = account.Password;

            //save and close excel file
            await Task.Run(() => excelPackage.SaveAs("ClientInfo.xlsx",
                SetPassword("PasswordClient.xlsx", "password")));
        }

        /// <summary>
        /// Bank withdrawal method. Change sum in excel file
        /// </summary>
        /// <param name="account">Account where the money is withdrawn from</param>
        /// <param name="amountMoney">Available sum on an account</param>
        /// <param name="tempAmountMoney">Desire sum to withdraw</param>
        /// <returns></returns>
        public static double WithdrawMoneyXLSX(IAccount account, double amountMoney, double tempAmountMoney)
        {
            int tempRow = 0;
            byte[] bin = File.ReadAllBytes("ClientInfo.xlsx");
            MemoryStream memoryStream = new MemoryStream(bin);
            ExcelPackage package1 = new ExcelPackage(memoryStream, GetPassword("PasswordClient.xlsx", "password"));
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
            package1.SaveAs("ClientInfo.xlsx", SetPassword("PasswordClient.xlsx", "password"));

            return amountMoney;
        }
        /// <summary>
        /// (Async)Bank withdrawal method. Change sum in excel file
        /// </summary>
        /// <param name="account">Account where the money is withdrawn from</param>
        /// <param name="amountMoney">Available sum on an account</param>
        /// <param name="tempAmountMoney">Desire sum to withdraw</param>
        /// <returns></returns>
        public static async Task<double> WithdrawMoneyXLSXAsync
            (IAccount account, double amountMoney, double tempAmountMoney)
        {
            int tempRow = 0;
            byte[] bin = File.ReadAllBytes("ClientInfo.xlsx");
            MemoryStream memoryStream = new MemoryStream(bin);
            ExcelPackage package1 = new ExcelPackage(memoryStream, GetPassword("PasswordClient.xlsx", "password"));
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
            await Task.Run(() => package1.SaveAs("ClientInfo.xlsx", SetPassword("PasswordClient.xlsx", "password")));
            return amountMoney;
        }

        /// <summary>
        /// (Async)ATM withdrawal method. Change sum in excel file
        /// </summary>
        /// <param name="atm">ATM where the money is withdrawn from</param>
        /// <param name="amountMoney">Available sum on an ATM</param>
        /// <param name="tempAmountMoney">Desire sum to withdraw</param>
        /// <returns></returns>
        public static async Task<double> WithdrawMoneyAtmXLSXAsync
            (IATM atm, double amountMoney, double tempAmountMoney)
        {
            int tempRow = 0;
            byte[] bin = File.ReadAllBytes("ATMInfo.xlsx");
            MemoryStream memoryStream = new MemoryStream(bin);
            ExcelPackage package1 = new ExcelPackage(memoryStream, GetPassword("PasswordATM.xlsx", "password"));
            ExcelWorksheet worksheet = package1.Workbook.Worksheets["ATM Info"];

            for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
            {
                for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.Start.Column; j++)
                {
                    int tempNumberATM = int.Parse(worksheet.Cells[i, j].Value.ToString());
                    if (atm.NumberATM == tempNumberATM)
                    {
                        tempRow = i;
                    }
                }
            }

            worksheet.Cells[tempRow, 3].Value = (amountMoney -= tempAmountMoney);
            await Task.Run(() => package1.SaveAs("ATMInfo.xlsx", GetPassword("PasswordATM.xlsx", "password")));
            return amountMoney;
        }
        /// <summary>
        /// ATM withdrawal method. Change sum in excel file
        /// </summary>
        /// <param name="atm">ATM where the money is withdrawn from</param>
        /// <param name="amountMoney">Available sum on an ATM</param>
        /// <param name="tempAmountMoney">Desire sum to withdraw</param>
        /// <returns></returns>
        public static double WithdrawMoneyAtmXLSX(IATM atm, double amountMoney, double tempAmountMoney)
        {
            int tempRow = 0;
            byte[] bin = File.ReadAllBytes("ATMInfo.xlsx");
            MemoryStream memoryStream = new MemoryStream(bin);
            ExcelPackage package1 = new ExcelPackage(memoryStream, GetPassword("PasswordATM.xlsx", "password"));
            ExcelWorksheet worksheet = package1.Workbook.Worksheets["ATM Info"];

            for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
            {
                for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.Start.Column; j++)
                {
                    int tempNumberATM = int.Parse(worksheet.Cells[i, j].Value.ToString());
                    if (atm.NumberATM == tempNumberATM)
                    {
                        tempRow = i;
                    }
                }
            }

            worksheet.Cells[tempRow, 3].Value = (amountMoney -= tempAmountMoney);
            package1.SaveAs("ATMInfo.xlsx", GetPassword("PasswordATM.xlsx", "password"));
            return amountMoney;
        }

        /// <summary>
        /// Create or open workbook with ATM info
        /// </summary>
        /// <param name="atm"></param>
        public static void WorksheetAtmXLSX(IATM atm)
        {
            ExcelPackage packageATM = new ExcelPackage();
            packageATM.Workbook.Properties.Author = "VGTAx";
            packageATM.Workbook.Properties.Company = "PVG";
            packageATM.Workbook.Properties.Title = "Information about ATM";
            packageATM.Workbook.Properties.Created = DateTime.Now;

            FileInfo fileInfo = new FileInfo("ATMInfo.xlsx");
            if (fileInfo.Exists)
            {
                packageATM = new ExcelPackage(fileInfo, GetPassword("PasswordATM.xlsx", "password"));
            }

            ExcelWorksheet? worksheetATM = packageATM.Workbook.Worksheets["ATM Info"];

            int rowWS = 0;
            int colWS = 0;

            if (worksheetATM != null)
            {
                rowWS = worksheetATM.Dimension.End.Row + 1;
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
            //check Number ATM for uniqueness
            if (rowWS > 2)
            {

                while (CheckInfoXLSX(atm.NumberATM, "ATMInfo.xlsx", "ATM Info", "PasswordATM.xlsx", "password") == false)
                {
                    atm.NumberATM = new Random().Next(1, 9999);
                }
            }
            //table(font and border) settings
            worksheetATM.Cells[1, 1, 1, colWS].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            worksheetATM.Cells[1, 1, 1, colWS].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            worksheetATM.Cells[1, 1, 1, colWS].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            worksheetATM.Cells[1, 1, rowWS, colWS].Style.Font.Size = 12;

            worksheetATM.Cells[1, (int)EnumATM.Number, rowWS, (int)EnumATM.Number].AutoFitColumns(8);
            worksheetATM.Cells[1, (int)EnumATM.Adress, rowWS, (int)EnumATM.Adress].AutoFitColumns(32);
            worksheetATM.Cells[1, (int)EnumATM.MoneyATM, rowWS, (int)EnumATM.MoneyATM].AutoFitColumns(15);

            worksheetATM.Columns[1, colWS].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheetATM.Columns[1, colWS].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheetATM.Columns[1, colWS].Style.WrapText = true;

            worksheetATM.Cells[2, 1, rowWS, colWS].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            worksheetATM.Cells[2, 1, rowWS, colWS].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            worksheetATM.Cells[2, 1, rowWS, colWS].Style.Border.Left.Style = ExcelBorderStyle.Medium;

            worksheetATM.Cells[2, (int)EnumATM.MoneyATM, rowWS, (int)EnumATM.MoneyATM].Style.Font.Bold = true;
            worksheetATM.Cells[2, (int)EnumATM.MoneyATM, rowWS, (int)EnumATM.MoneyATM].Style.Font.Color.SetColor(Color.Red);

            worksheetATM.Cells[rowWS, (int)EnumATM.Number].Value = atm.NumberATM;
            worksheetATM.Cells[rowWS, (int)EnumATM.Adress].Value = atm.Adress;
            worksheetATM.Cells[rowWS, (int)EnumATM.MoneyATM].Value = atm.AmountOfMoneyATM;

            packageATM.SaveAs("ATMInfo.xlsx", SetPassword("PasswordATM.xlsx", "password"));

        }
        /// <summary>
        /// (Async)Create or open workbook with ATM info
        /// </summary>
        /// <param name="atm"></param>
        public static async Task WorksheetAtmXLSXAsync(IATM atm)
        {
            ExcelPackage packageATM = new ExcelPackage();
            packageATM.Workbook.Properties.Author = "VGTAx";
            packageATM.Workbook.Properties.Company = "PVG";
            packageATM.Workbook.Properties.Title = "Information about ATM";
            packageATM.Workbook.Properties.Created = DateTime.Now;

            FileInfo fileInfo = new FileInfo("ATMInfo.xlsx");
            if (fileInfo.Exists)
            {
                packageATM = new ExcelPackage(fileInfo, GetPassword("PasswordATM.xlsx", "password"));
            }

            ExcelWorksheet? worksheetATM = packageATM.Workbook.Worksheets["ATM Info"];

            int rowWS = 0;
            int colWS = 0;

            if (worksheetATM != null)
            {
                rowWS = worksheetATM.Dimension.End.Row + 1;
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
            //check Number ATM for uniqueness
            if (rowWS > 2)
            {

                while (CheckInfoXLSX(atm.NumberATM, "ATMInfo.xlsx", "ATM Info", "PasswordATM.xlsx", "password") == false)
                {
                    atm.NumberATM = new Random().Next(1, 9999);
                }
            }
            //table(font and border) settings
            worksheetATM.Cells[1, 1, 1, colWS].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            worksheetATM.Cells[1, 1, 1, colWS].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            worksheetATM.Cells[1, 1, 1, colWS].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            worksheetATM.Cells[1, 1, rowWS, colWS].Style.Font.Size = 12;

            worksheetATM.Cells[1, (int)EnumATM.Number, rowWS, (int)EnumATM.Number].AutoFitColumns(8);
            worksheetATM.Cells[1, (int)EnumATM.Adress, rowWS, (int)EnumATM.Adress].AutoFitColumns(32);
            worksheetATM.Cells[1, (int)EnumATM.MoneyATM, rowWS, (int)EnumATM.MoneyATM].AutoFitColumns(15);

            worksheetATM.Columns[1, colWS].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            worksheetATM.Columns[1, colWS].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheetATM.Columns[1, colWS].Style.WrapText = true;

            worksheetATM.Cells[2, 1, rowWS, colWS].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            worksheetATM.Cells[2, 1, rowWS, colWS].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            worksheetATM.Cells[2, 1, rowWS, colWS].Style.Border.Left.Style = ExcelBorderStyle.Medium;

            worksheetATM.Cells[2, (int)EnumATM.MoneyATM, rowWS, (int)EnumATM.MoneyATM].Style.Font.Bold = true;
            worksheetATM.Cells[2, (int)EnumATM.MoneyATM, rowWS, (int)EnumATM.MoneyATM].Style.Font.Color.SetColor(Color.Red);

            worksheetATM.Cells[rowWS, (int)EnumATM.Number].Value = atm.NumberATM;
            worksheetATM.Cells[rowWS, (int)EnumATM.Adress].Value = atm.Adress;
            worksheetATM.Cells[rowWS, (int)EnumATM.MoneyATM].Value = atm.AmountOfMoneyATM;

            await Task.Run(() => packageATM.SaveAs("ATMInfo.xlsx", SetPassword("PasswordATM.xlsx", "password")));
        }

        //Load list accounts from Excel file
        public static List<IAccount> LoadListAccXLSX()
        {
            List<IAccount> accountsList = new List<IAccount>();
            try
            {
                byte[]? bin = File.ReadAllBytes("ClientInfo.xlsx");
                MemoryStream memoryStream = new MemoryStream(bin);
                ExcelPackage excelPackage = new ExcelPackage(memoryStream, GetPassword("PasswordClient.xlsx", "password"));
                ExcelWorksheet clientInfo = excelPackage.Workbook.Worksheets["ClientInfo"];
                ExcelWorksheet accInfo = excelPackage.Workbook.Worksheets["ClientAccountInfo"];

                for (int i = clientInfo.Dimension.Start.Row + 1; i <= clientInfo.Dimension.End.Row; i++)
                {
                    Account temp = new Account();

                    //write properties from file to object
                    temp.ID = int.Parse(clientInfo.Cells[i, (int)EnumClient.ID].Value.ToString());
                    temp.person.Name = clientInfo.Cells[i, (int)EnumClient.Name].Value.ToString();
                    temp.person.SurName = clientInfo.Cells[i, (int)EnumClient.Surname].Value.ToString();
                    temp.person.Age = int.Parse(clientInfo.Cells[i, (int)EnumClient.Age].Value.ToString());
                    temp.AmountOfMoney = int.Parse(clientInfo.Cells[i, (int)EnumClient.Money].Value.ToString());

                    temp.Login = accInfo.Cells[i, (int)EnumAcc.Login].Value.ToString();
                    temp.Password = accInfo.Cells[i, (int)EnumAcc.Password].Value.ToString();

                    // add object to list
                    accountsList.Add(temp);
                }
                // return list
                return accountsList;
            }
            catch (Exception)
            {
                return accountsList;
            }
        }
        //Load list ATM from Excel file
        public static List<IATM> LoadListAtmXLSX()
        {
            List<IATM> listATM = new List<IATM>();
            try
            {
                byte[] bin = File.ReadAllBytes("ATMInfo.xlsx");
                MemoryStream memoryStream = new MemoryStream(bin);
                ExcelPackage packageATM = new ExcelPackage(memoryStream, GetPassword("PasswordATM.xlsx", "password"));
                ExcelWorksheet worksheet = packageATM.Workbook.Worksheets["ATM Info"];

                for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
                {
                    ATM atm = new ATM();

                    atm.NumberATM = int.Parse(worksheet.Cells[i, (int)EnumATM.Number].Value.ToString());
                    atm.Adress = worksheet.Cells[i, (int)EnumATM.Adress].Value.ToString();
                    atm.AmountOfMoneyATM = int.Parse(worksheet.Cells[i, (int)EnumATM.MoneyATM].Value.ToString());

                    //add object to list
                    listATM.Add(atm);
                }
                ///retur list ATM
                return listATM;
            }
            catch (Exception)
            {
                return listATM;
            }
        }

        /// <summary>
        /// getting a password to access an excel file
        /// </summary>
        /// <param name="workbook">Name workbook with .xlsx format</param>
        /// <param name="worksheet">Name worksheet in the workbook</param>
        /// <returns></returns>
        public static string GetPassword(string workbook, string worksheet)
        {
            byte[]? bin1 = File.ReadAllBytes(workbook);
            MemoryStream memoryStream1 = new MemoryStream(bin1);
            ExcelPackage excel = new ExcelPackage(memoryStream1);
            ExcelWorksheet excelWorksheet = excel.Workbook.Worksheets[worksheet];

            return excelWorksheet.Cells[2, 1].Value.ToString();
        }

        /// <summary>
        /// setting a password to access an excel file
        /// </summary>
        /// <param name="workbook">Name workbook with .xlsx format</param>
        /// <param name="worksheet">Name worksheet in the workbook</param>
        /// <returns></returns>
        public static string SetPassword(string workbook, string worksheet)
        {
            FileInfo filePass = new FileInfo(workbook);
            ExcelPackage? passPack = new ExcelPackage(filePass);
            ExcelWorksheet? passWS = passPack.Workbook.Worksheets[worksheet];

            if (passWS == null)
            {
                passWS = passPack.Workbook.Worksheets.Add(worksheet);

                var pass = passWS.Cells["A1"];
                var title = pass.RichText.Add("Password");
            }
            //generate password
            string passwordWB = new Random().Next(10000000, 99999999).ToString();
            string plain = "Test";
            //Encrypt password
            string encryptPassword = Rijndael.Encrypt(plain, passwordWB, KeySize.Aes256);

            passWS.Protection.SetPassword(passwordWB);
            passWS.Cells[2, 1].Value = encryptPassword;
            passPack.SaveAs(workbook);
            return encryptPassword;
        }
    }
}
