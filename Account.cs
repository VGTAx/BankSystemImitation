using InitHelperInformatMessage;
using Attributes;
using OfficeOpenXml.Style;
using OfficeOpenXml;
using System.Drawing;

namespace BankSystem
{
    /// <summary>
    /// сlass Account used to create a bank customer
    /// </summary>
    
    [CheckLengthLoginPassword(15)]  
    internal class Account : IAccount
    {        
        public event Func<string, double> ?doubleMethod;
        public event Func<string, string> ?stringMethod;
        public event Func<string,int> ?intMethod;

        public Account()
        {
            personObj = new Person();
        }
        public Account(Person person, string password, string login, double sum, int iD)
        {            
            ID = iD;
            personObj = person;                      
            Login = login;
            Password = password;
            AmountOfMoney = sum;
        }
    
        public Person personObj { get; set; }
        public string Password { get; set; } 
        public string Login { get; set; }
        public int ID { get; set; }
        public string Surname { get; set; }
        public double AmountOfMoney { get; set; }
        public bool Authorization { get; set; }

        [Obsolete]
        public bool LoginAcc()
        {
            string login = InitializationHelper.StringInIt("Login: ").ToUpper();
            string password = InitializationHelper.StringInIt("Password: ").ToUpper();

            int count = 3;
            while (login != Login || password != Password && count > 0)
            {
                MessageInformant.ErrorOutput($"Invalid login or password! {count} attemp left");

                login = InitializationHelper.StringInIt("Login: ").ToUpper();
                password = InitializationHelper.StringInIt("Password: ").ToUpper();

                count--;
                if (count == 0)
                {
                    MessageInformant.ErrorOutput("attempt limit reached");
                    return false;
                }
            }
            MessageInformant.SuccessOutput("Login succeeded");
            return true;
        }

        /// <summary>
        /// Displays/returns information about the client
        /// </summary>        /// 
        /// <returns></returns>
        public string Info()
        {
            Console.Clear();
            Console.WriteLine($"Account Info:\n");
            personObj.Info();
            Console.WriteLine($"Amount of money - {AmountOfMoney} BYN");
            Console.ReadLine();

            return $"{personObj.Info}\nAmount of money - {AmountOfMoney} BYN";
        }
       
        /// <summary>
        /// Register account customer
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public Account RegistrAcc(List<IAccount> obj)
        {
            stringMethod += InitializationHelper.StringInIt;
            intMethod += InitializationHelper.IntInit;
            do
            {
                Login = stringMethod.Invoke("login account");

                //login availability check
                foreach (var itLog in obj)
                {
                    while (Login == itLog.Login)
                    {
                        MessageInformant.ErrorOutput($"Login \"{Login}\" not available");
                        Login = stringMethod.Invoke("login acc");
                    }
                }

                Password = stringMethod.Invoke("password acc");
                //password matching check
                while (stringMethod?.Invoke("repeat pass") != Password)
                {
                    MessageInformant.ErrorOutput("Passwords do not match");
                }

                string namePerson = stringMethod.Invoke("name");
                string surnamePerson = stringMethod.Invoke("surname");
                int agePerson = intMethod.Invoke("age");
                Random rand = new Random(); 
                ID = rand.Next(1,999);                
                personObj = new Person(agePerson,namePerson,surnamePerson);
                while(!Person.CheckAge(personObj))
                {                    
                    agePerson = intMethod.Invoke("age");
                    personObj = new Person(agePerson, namePerson, surnamePerson);
                }
            }
            while (CheckLength(new Account(personObj, Password, Login, AmountOfMoney,ID))==false);


            
            MessageInformant.SuccessOutput("Account registered!");
            TestExcel(new Account(personObj, Password, Login, AmountOfMoney, ID));
            return new Account(personObj, Password, Login, AmountOfMoney, ID);
        }
       
        /// <summary>
        /// Deposit money into an account
        /// </summary>
        /// <param name="amount"></param>
        /// <returns></returns>
        public double AddMoney(double amount = 0)
        {
            doubleMethod += InitializationHelper.DoubleInit;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;  
            ExcelPackage excelPackage = new ExcelPackage();
            FileInfo fileInfo = new FileInfo("ClientInfo.xlsx");   

            if (fileInfo.Exists)
            {
                excelPackage = new ExcelPackage(fileInfo);
                ExcelWorksheet? ClientInfoWS = excelPackage.Workbook.Worksheets["ClientInfo"];
                int rowClient = 1;

                for (int i = ClientInfoWS.Dimension.Start.Row; i <= ClientInfoWS.Dimension.End.Row; i++)
                {
                    rowClient++;
                }

                if (amount == 0)
                {
                    ClientInfoWS.Cells[rowClient, 5].Value = AmountOfMoney += doubleMethod.Invoke("sum of money to add"); 
                    excelPackage.SaveAs("ClientInfo.xlsx");
                    return AmountOfMoney;
                }
                else
                {
                    ClientInfoWS.Cells[rowClient, 5].Value = AmountOfMoney += amount; 
                    excelPackage.SaveAs("ClientInfo.xlsx");
                    return AmountOfMoney;
                }
                    
            }
            return AmountOfMoney;
        }
       
        /// <summary>
        /// Withdraw money from the account
        /// </summary>
        /// <param name="desiredAmount"></param>
        /// <returns></returns>
        public double TakeMoney(double desiredAmount = 0)
        {
            //money check
            if (desiredAmount == 0)
            {
                double temp = doubleMethod.Invoke("sum of money to withdraw");

                while (temp == 0 || temp > AmountOfMoney)
                {
                    if (temp == 0)//while try to withdraw 0 BYN
                    {
                        MessageInformant.ErrorOutput($"Can't withdraw 0 BYN");
                        temp = doubleMethod.Invoke("sum of money to withdraw");
                    }
                    else//while not enough money
                    {
                        MessageInformant.ErrorOutput($"Not enough money. You have {AmountOfMoney} BYN");
                        temp = doubleMethod.Invoke("sum of money to withdraw");
                    }
                }
                MessageInformant.SuccessOutput($"Money withdrawn {temp} BYN");
                return AmountOfMoney -= temp;
            }
            //money check
            if (desiredAmount > AmountOfMoney)
            {
                MessageInformant.ErrorOutput($"Not enough money. You have {AmountOfMoney} BYN");
                return -1;
            }
            else
                return AmountOfMoney -= desiredAmount;
        }

        private static bool CheckLength(Account account)
        {
            Type? type = typeof(Account);
            object[] attributes = type.GetCustomAttributes(false);
            foreach (Attribute attr in attributes)
            {
                if (attr is  CheckLengthLoginPasswordAttribute attribute)
                {
                    if (attribute.Length > account.Login.Length)
                    {
                        if (attribute.Length > account.Password.Length)
                            return true;
                        MessageInformant.ErrorOutput($"Length Login must be less {attribute.Length}");
                        return false;
                    }
                    else
                    {
                        MessageInformant.ErrorOutput($"Length Password must be less {attribute.Length}");
                        return false;
                    }
                }
            }
            return true;
        }

        public void TestExcel(IAccount account)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excelPackage = new ExcelPackage();
            excelPackage.Workbook.Properties.Author = "VGTAx";
            excelPackage.Workbook.Properties.Company = "PVG";
            excelPackage.Workbook.Properties.Title = "Title";
            excelPackage.Workbook.Properties.Created = DateTime.Now;
           
            FileInfo fileInfo = new FileInfo("ClientInfo.xlsx");
            if(fileInfo.Exists)
            {
                excelPackage = new ExcelPackage(fileInfo);
            }      


            ExcelWorksheet? ClientInfoWS = excelPackage.Workbook.Worksheets["ClientInfo"];
            ExcelWorksheet? ClientAccInfoWS = excelPackage.Workbook.Worksheets["ClientAccountInfo"];


            if (ClientInfoWS == null && ClientAccInfoWS == null)
            {                
                ClientInfoWS = excelPackage.Workbook.Worksheets.Add("ClientInfo");
                var ID = ClientInfoWS.Cells[1, 1];
                var name = ClientInfoWS.Cells["B1:E1"];
                var surname = ClientInfoWS.Cells[1,3];
                var age = ClientInfoWS.Cells[1,4];
                var AmountOfMoney = ClientInfoWS.Cells[1,5];
                

                ClientAccInfoWS = excelPackage.Workbook.Worksheets.Add("ClientAccountInfo");
                var login = ClientAccInfoWS.Cells["B1:C1"];
                var password = ClientAccInfoWS.Cells[1, 3];

                name.IsRichText = true;
                name.Style.WrapText = true;
                
                var borderName = name.Style.Border.Bottom.Style = name.Style.Border.Right.Style = name.Style.Border.Left.Style = name.Style.Border.Top.Style = ExcelBorderStyle.Medium;

                var titleID = ID.RichText.Add("ID");
                var titleName = name.RichText.Add("Name");
                var titleSurname = surname.RichText.Add("Surname");
                var titleAge = age.RichText.Add("Age");
                var titleAmountOfMoney = AmountOfMoney.RichText.Add("Amount of Money");
                var titleLogin = login.RichText.Add("Login");
                var titlePassword = password.RichText.Add("Password");  

                titleName.Bold = true;
                titleName.FontName = "Cambria";
                titleName.Size = 14;                

                List<ExcelRichText> list = new List<ExcelRichText>();
                list.Add(titleID);
                list.Add(titleAge);
                list.Add(titleSurname);
                list.Add(titleAmountOfMoney);
                list.Add(titleLogin);
                list.Add(titlePassword);

                foreach (var item in list)
                {
                    item.Bold = titleName.Bold;
                    item.FontName = titleName.FontName;
                    item.Size = 14;                      
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
           
            if(rowClient > 2)
            {
                checkID(account.ID);
                while (checkID(account.ID) == false)
                {
                    ID = new Random().Next(1, 9999);
                }
            }

            ClientInfoWS.Cells[rowClient, 1].Value = account.ID;
            ClientInfoWS.Cells[rowClient, 2].Value = account.personObj.Name;
            ClientInfoWS.Cells[rowClient, 3].Value = account.personObj.SurName;
            ClientInfoWS.Cells[rowClient, 4].Value = account.personObj.Age;            
            ClientAccInfoWS.Cells[rowAccount, 2].Value = account.Login;
            ClientAccInfoWS.Cells[rowAccount, 3].Value = account.Password;


            excelPackage.SaveAs("ClientInfo.xlsx");
        }


        public bool checkID(int newID)
        {

            byte[] bin = File.ReadAllBytes("ClientInfo.xlsx");
            MemoryStream memoryStream = new MemoryStream(bin);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage excelPackage = new ExcelPackage();
            
            if(memoryStream.CanRead)
            {
                excelPackage = new ExcelPackage(memoryStream);
                ExcelWorksheet? ClientInfoWS = excelPackage.Workbook.Worksheets["ClientInfo"]; 
                for(int i = ClientInfoWS.Dimension.Start.Row; i <= ClientInfoWS.Dimension.Start.Row; i++)
                {
                    for (int j = ClientInfoWS.Dimension.Start.Column; j <= ClientInfoWS.Dimension.End.Column; j++)
                    {
                        string temp = ClientInfoWS.Cells[j, i].Value.ToString();
                        //if (temp == newID)                       
                                                  
                    }
                }                
            }            
            return true;
        }
        
        
    }

}

