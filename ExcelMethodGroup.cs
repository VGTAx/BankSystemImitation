using OfficeOpenXml;
using OfficeOpenXml.Style;
using InitHelperInformatMessage;

namespace BankSystem
{
    internal class ExcelMethodGroup
    {
        public static bool checkID(int newID)
        {
            try
            {
                byte[] bin = File.ReadAllBytes("ClientInfo.xlsx");
                MemoryStream memoryStream = new MemoryStream(bin);                
                ExcelPackage excelPackage = new ExcelPackage(memoryStream);
                ExcelWorksheet? ClientInfoWS = excelPackage.Workbook.Worksheets["ClientInfo"];

                for (int i = ClientInfoWS.Dimension.Start.Row + 1; i < ClientInfoWS.Dimension.End.Row; i++)
                {
                    for (int j = ClientInfoWS.Dimension.Start.Column; j <= ClientInfoWS.Dimension.Start.Column; j++)
                    {
                        int tempID = int.Parse(ClientInfoWS.Cells[i, j].Value.ToString());
                        if (tempID == newID)
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

        public static bool CheckAccNameAvailable(string login)
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

        public static void TestExcel(IAccount account)
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
                checkID(account.ID);
                while (checkID(account.ID) == false)
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


        public static double WithDrawMoney(IAccount account, double amountMoney, double tempAmountMoney)
        {
            int tempRow = 0;
            byte[] bin = File.ReadAllBytes("ClientInfo.xlsx");
            MemoryStream memoryStream = new MemoryStream(bin);
            ExcelPackage package1 = new ExcelPackage(memoryStream);
            ExcelWorksheet worksheet = package1.Workbook.Worksheets["ClientInfo"];

            for (int i = worksheet.Dimension.Start.Row + 1; i < worksheet.Dimension.End.Row; i++)
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
            MessageInformant.SuccessOutput($"Money withdrawn {tempAmountMoney} BYN");
            
            return amountMoney;
        }
        
    }
}
