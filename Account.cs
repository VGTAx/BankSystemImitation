using InitHelperInformatMessage;
using Attributes;
using OfficeOpenXml.Style;
using OfficeOpenXml;


namespace BankSystem
{
    /// <summary>
    /// сlass Account used to create a bank customer
    /// </summary>
    
    [CheckLengthLoginPassword(20)]  
    internal class Account : IAccount
    {        
        //public event Func<string, double> ?doubleMethod;
        //public event Func<string, string> ?stringMethod;
        //public event Func<string,int> ?intMethod;

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
            do
            {
                Login = InitializationHelper.StringInIt("login account");
                while(!ExcelMethodGroup.CheckAccAvailable(Login))
                {
                    Login = InitializationHelper.StringInIt("login account");
                }
                
                Password = InitializationHelper.StringInIt("password acc");
                //password matching check
                while (InitializationHelper.StringInIt("repeat pass") != Password)
                {
                    MessageInformant.ErrorOutput("Passwords do not match");
                }

                string namePerson = InitializationHelper.StringInIt("name");
                string surnamePerson = InitializationHelper.StringInIt("surname");
                int agePerson =InitializationHelper.IntInit("age");

                Random rand = new Random(); 
                ID = rand.Next(1,10000); 
                
                personObj = new Person(agePerson,namePerson,surnamePerson);
                while(!Person.CheckAge(personObj))
                {                    
                    agePerson =InitializationHelper.IntInit("age");
                    personObj = new Person(agePerson, namePerson, surnamePerson);
                }
            }
            while (CheckLength(new Account(personObj, Password, Login, 0,ID))==false);
            
            MessageInformant.SuccessOutput("Account registered!");
            ExcelMethodGroup.TestExcel(new Account(personObj, Password, Login, 0, ID));
            return new Account(personObj, Password, Login, 0, ID);
        }
       
        /// <summary>
        /// Deposit money into an account
        /// </summary>
        /// <param name="amount"></param>
        /// <returns></returns>
        public double AddMoney(double amount = 0)
        {    
            ExcelPackage excelPackage = new ExcelPackage();
            FileInfo fileInfo = new FileInfo("ClientInfo.xlsx");   

            if (fileInfo.Exists)
            {
                excelPackage = new ExcelPackage(fileInfo);
                ExcelWorksheet? ClientInfoWS = excelPackage.Workbook.Worksheets["ClientInfo"];
                int rowClient = 1;

                for (int i = ClientInfoWS.Dimension.Start.Row+1; i <= ClientInfoWS.Dimension.End.Row; i++)
                {
                    if(ID == int.Parse(ClientInfoWS.Cells[i,1].Value.ToString()))
                    {
                        rowClient = i;
                        break;
                    }                   
                }

                if (amount == 0)
                {
                    ClientInfoWS.Cells[rowClient, 5].Value = 
                        AmountOfMoney += InitializationHelper.DoubleInit("sum of money to add"); 
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
                double tempDesAmount = InitializationHelper.DoubleInit("sum of money to withdraw");

                while (tempDesAmount == 0 || tempDesAmount > AmountOfMoney)
                {
                    if (tempDesAmount == 0)//while try to withdraw 0 BYN
                    {
                        MessageInformant.ErrorOutput($"Can't withdraw 0 BYN");
                        tempDesAmount = InitializationHelper.DoubleInit("sum of money to withdraw");
                    }
                    else//while not enough money
                    {
                        MessageInformant.ErrorOutput($"Not enough money. You have {AmountOfMoney} BYN");
                        tempDesAmount = InitializationHelper.DoubleInit("sum of money to withdraw");
                    }
                }
                return AmountOfMoney = ExcelMethodGroup.WithDrawMoney(this, AmountOfMoney, tempDesAmount);                
            }
            //money check
            if (desiredAmount > AmountOfMoney)
            {
                MessageInformant.ErrorOutput($"Not enough money. You have {AmountOfMoney} BYN");
                return -1;
            }
            else
                return AmountOfMoney = ExcelMethodGroup.WithDrawMoney(this, AmountOfMoney, desiredAmount);
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
        
    }

}

