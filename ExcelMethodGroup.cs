using OfficeOpenXml;
using OfficeOpenXml.Style;
using InitHelperInformatMessage;
using System.Drawing;
using Rijndael256;

namespace BankSystem
{
    internal class ExcelMethodGroup
    {
        /// <summary>
        /// Сhecking for repetitions of the passed argument
        /// </summary>
        /// <param name="value">Сhecked value (ID or №ATM)</param>
        /// <param name="workbook">Name workbook with format(xlsx)</param>
        /// <param name="worksheets">Name worksheet in the workbook</param>
        /// <returns></returns>
        public static bool CheckInfoXLSX(int value, string workbook, string worksheets,string WBPass ="",
            string WSPass="")
        {
            try
            {
                ///open file for read
                byte[] bin = File.ReadAllBytes(workbook);
                MemoryStream memoryStream = new MemoryStream(bin);                
                ExcelPackage excelPackage = new ExcelPackage(memoryStream,GetPassword(WBPass, WSPass));
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
        /// Сhecking for repetitions Account name (Login)
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
                var name = ClientInfoWS.Cells[1,2];
                var surname = ClientInfoWS.Cells[1, 3];
                var age = ClientInfoWS.Cells[1, 4];
                var AmountOfMoney = ClientInfoWS.Cells[1, 5];
                //create worksheet "Account Info" (login & password)
                ClientAccInfoWS = excelPackage.Workbook.Worksheets.Add("ClientAccountInfo");
                var login = ClientAccInfoWS.Cells["B1:C1"];
                var password = ClientAccInfoWS.Cells[1, 3];

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
           
            int rowAccountWS = ClientAccInfoWS.Dimension.End.Row+1;
            int colAccountWS = ClientAccInfoWS.Columns.EndColumn;            
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
            ClientInfoWS.Cells[1,1,1,colClientWS].Style.Border.Right.Style = ExcelBorderStyle.Medium;
            ClientInfoWS.Cells[1,1,1,colClientWS].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            ClientInfoWS.Cells[1,1,1,colClientWS].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            ClientInfoWS.Cells[1,1,rowClientWS, colClientWS].Style.Font.Size = 12;

            ClientInfoWS.Cells[1,1,rowClientWS,1].AutoFitColumns(6);
            ClientInfoWS.Cells[1,2,rowClientWS,2].AutoFitColumns(20);
            ClientInfoWS.Cells[1,3,rowClientWS,3].AutoFitColumns(20);
            ClientInfoWS.Cells[1,4,rowClientWS,4].AutoFitColumns(6);
            ClientInfoWS.Cells[1,5,rowClientWS,5].AutoFitColumns(10);

            ClientInfoWS.Columns[1,colClientWS].Style.VerticalAlignment= ExcelVerticalAlignment.Center;
            ClientInfoWS.Columns[1,colClientWS].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            ClientInfoWS.Columns[1,colClientWS].Style.WrapText = true;
            
            ClientInfoWS.Cells[2, 1, rowClientWS, colClientWS].Style.Border.Bottom.Style = ExcelBorderStyle.Medium;
            ClientInfoWS.Cells[2, 1, rowClientWS, colClientWS].Style.Border.Left.Style = ExcelBorderStyle.Medium;
            ClientInfoWS.Cells[2, 1, rowClientWS, colClientWS].Style.Border.Right.Style = ExcelBorderStyle.Medium;

            ClientInfoWS.Cells[2, 5, rowClientWS, 5].Style.Font.Color.SetColor(Color.Red);
            ClientInfoWS.Cells[2, 5, rowClientWS, 5].Style.Font.Bold = true;

            ClientInfoWS.Cells[rowClientWS, 1].Value = account.ID;
            ClientInfoWS.Cells[rowClientWS, 2].Value = account.personObj.Name;
            ClientInfoWS.Cells[rowClientWS, 3].Value = account.personObj.SurName;
            ClientInfoWS.Cells[rowClientWS, 4].Value = account.personObj.Age;
            ClientInfoWS.Cells[rowClientWS, 5].Value = account.AmountOfMoney;

            ClientAccInfoWS.Cells[rowAccountWS, 2].Value = account.Login;
            ClientAccInfoWS.Cells[rowAccountWS, 3].Value = account.Password;
            //save and close excel file
            excelPackage.SaveAs("ClientInfo.xlsx", SetPassword("PasswordClient.xlsx", "password"));           
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
            ExcelPackage package1 = new ExcelPackage(memoryStream,GetPassword("PasswordClient.xlsx", "password"));
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
            package1.SaveAs("ClientInfo.xlsx",SetPassword("PasswordClient.xlsx", "password"));
            return amountMoney;
        }

        /// <summary>
        /// ATM withdrawal method. Change sum in excel file
        /// </summary>
        /// <param name="bankomat">Account where the money is withdrawn from</param>
        /// <param name="amountMoney">Available sum on an ATM</param>
        /// <param name="tempAmountMoney">Desire sum to withdraw</param>
        /// <returns></returns>
        public static double WithdrawMoneyAtmXLSX(IBankomat bankomat, double amountMoney, double tempAmountMoney)
        {
            int tempRow = 0;
            byte[] bin = File.ReadAllBytes("ATMInfo.xlsx");
            MemoryStream memoryStream = new MemoryStream(bin);
            ExcelPackage package1 = new ExcelPackage(memoryStream,GetPassword("PasswordATM.xlsx", "password"));
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
            package1.SaveAs("ATMInfo.xlsx",GetPassword("PasswordATM.xlsx", "password"));            
            return amountMoney;
        }

        /// <summary>
        /// Create or open workbook with ATM info
        /// </summary>
        /// <param name="bankomat"></param>
        public static void WorksheetAtmXLSX(IBankomat bankomat)
        {
            ExcelPackage packageATM = new ExcelPackage();
            packageATM.Workbook.Properties.Author = "VGTAx";
            packageATM.Workbook.Properties.Company = "PVG";
            packageATM.Workbook.Properties.Title = "Information about ATM";
            packageATM.Workbook.Properties.Created = DateTime.Now;

            FileInfo fileInfo = new FileInfo("ATMInfo.xlsx");
            if (fileInfo.Exists)
            {
                packageATM = new ExcelPackage(fileInfo,GetPassword("PasswordATM.xlsx", "password"));
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
            //check Number ATM for uniqueness
            if (rowWS > 2)
            {
               
                while (CheckInfoXLSX(bankomat.NumberATM, "ATMInfo.xlsx", "ATM Info", "PasswordATM.xlsx", "password") == false)
                {
                    bankomat.NumberATM = new Random().Next(1, 9999);
                }
            }
            //table(font and border) settings
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
            
            packageATM.SaveAs("ATMInfo.xlsx", SetPassword("PasswordATM.xlsx", "password"));
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
                    for (int j = clientInfo.Dimension.Start.Column; j < clientInfo.Dimension.End.Column + 1; j++)
                    {
                        //property:ID
                        if (j == 1)
                        {
                            temp.ID = int.Parse(clientInfo.Cells[i, j].Value.ToString());
                        }
                        //properties:Name and Login
                        if (j == 2)
                        {
                            temp.personObj.Name = clientInfo.Cells[i, j].Value.ToString();
                            temp.Login = accInfo.Cells[i, j].Value.ToString();
                        }
                        //properties:Surnmae and Password
                        if (j == 3)
                        {
                            temp.personObj.SurName = clientInfo.Cells[i, j].Value.ToString();
                            temp.Password = accInfo.Cells[i, j].Value.ToString();
                        }
                        //properties:Age
                        if (j == 4)
                            temp.personObj.Age = int.Parse(clientInfo.Cells[i, j].Value.ToString());
                        //properties:Amount of money
                        if (j == 5)
                        {
                            temp.AmountOfMoney = int.Parse(clientInfo.Cells[i, j]?.Value.ToString());
                        }

                    }
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
        public static List<IBankomat> LoadListAtmXLSX()
        {
            List<IBankomat> listATM = new List<IBankomat>();
            try
            {
                byte[] bin = File.ReadAllBytes("ATMInfo.xlsx");
                MemoryStream memoryStream = new MemoryStream(bin);
                ExcelPackage packageATM = new ExcelPackage(memoryStream, GetPassword("PasswordATM.xlsx", "password"));
                ExcelWorksheet worksheet = packageATM.Workbook.Worksheets["ATM Info"];

                for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
                {
                    Bankomat bankomat = new Bankomat();
                    //write properties from file to object
                    for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                    {
                        //Property:Number ATM
                        if (j == 1)
                            bankomat.NumberATM = int.Parse(worksheet.Cells[i, j].Value.ToString());
                        //Property:Adress ATM
                        if (j == 2)
                            bankomat.Adress = worksheet.Cells[i, j].Value.ToString();
                        //Property:Amount of money ATM
                        if (j == 3)
                            bankomat.AmountOfMoneyATM = int.Parse(worksheet.Cells[i, j].Value.ToString());
                    }
                    //add object to list
                    listATM.Add(bankomat);
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

            return  excelWorksheet.Cells[2, 1].Value.ToString();
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
            ExcelPackage ?passPack = new ExcelPackage(filePass);
            ExcelWorksheet? passWS = passPack.Workbook.Worksheets[worksheet];
            
            if (passWS == null)
            {
                passWS = passPack.Workbook.Worksheets.Add(worksheet);
                
                var pass = passWS.Cells["A1"];
                var title = pass.RichText.Add("Password");
            }
            //generate password
            string passwordWB = new Random().Next(1000000, 9999999).ToString();
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
