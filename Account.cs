using Attributes;
using InitHelperInformatMessage;
using OfficeOpenXml;
using System.Xml.Linq;

namespace BankSystem
{
    /// <summary>
    /// сlass Account used to create a bank customer
    /// </summary>

    [CheckLengthLoginPassword(15)]
    public class Account : IAccount
    {
        public Account()
        {
            person = new Person();
        }
        public Account(Person person, string password, string login, double sum, int iD)
        {
            ID = iD;
            this.person = person;
            Login = login;
            Password = password;
            AmountOfMoney = sum;
        }

        public Person person { get; set; }
        public string Password { get; set; }
        public string Login { get; set; }
        public int ID { get; set; }
        public string Surname { get; set; }
        public double AmountOfMoney { get; set; }
        public bool Authorization { get; set; }


        /// <summary>
        /// Displays/returns information about the client
        /// </summary>        /// 
        /// <returns></returns>
        public string Info()
        {
            Console.Clear();
            Console.WriteLine($"Account Info:\n");
            person.Info();
            Console.WriteLine($"Amount of money - {AmountOfMoney} BYN");
            Console.ReadLine();

            return $"{person.Info}\nAmount of money - {AmountOfMoney} BYN";
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
                while (!XmlMethodGroup.CheckAccNameAvailableXml(Login))
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
                int agePerson = InitializationHelper.IntInit("age");

                Random rand = new Random();
                ID = rand.Next(1, 10000);
                ///check if ID is available
                while (!XmlMethodGroup.CheckValueAvailableXml(ID, XmlMethodGroup.ACCOUNTS_INFO_DOCUMENT,
                                                    XmlMethodGroup.ACCOUNTS_INFO_ELEM, XmlMethodGroup.ID))
                {
                    ID = rand.Next(1, 10000);
                }

                person = new Person(agePerson, namePerson, surnamePerson);
                while (!Person.CheckAge(person))
                {
                    agePerson = InitializationHelper.IntInit("age");
                    person.Age = agePerson;
                }
            }
            while (CheckLengthLoginAndPassword(new Account(person, Password, Login, 0, ID)) == false);
            ///Excel file like a database
            //ExcelMethodGroup.WorksheetAccountXLSXAsync(new Account(person, Password, Login, 0, ID));

            //Xml file like a databse
            XmlMethodGroup.OpenOrCreateXmlAccountFile(new Account(person, Password, Login, 0, ID));
            MessageInformant.SuccessOutput("Account registered!");
            Console.ReadLine();
            Console.Clear();
            return new Account(person, Password, Login, 0, ID);
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
                excelPackage = new ExcelPackage(fileInfo, ExcelMethodGroup.GetPassword("PasswordClient.xlsx", "password"));
                ExcelWorksheet? ClientInfoWS = excelPackage.Workbook.Worksheets["ClientInfo"];
                int rowClient = 1;

                for (int i = ClientInfoWS.Dimension.Start.Row + 1; i <= ClientInfoWS.Dimension.End.Row; i++)
                {
                    if (ID == int.Parse(ClientInfoWS.Cells[i, 1].Value.ToString()))
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
                    MessageInformant.SuccessOutput($"Money added!");
                    Console.ReadLine();
                    return AmountOfMoney;
                }
                else
                {
                    ClientInfoWS.Cells[rowClient, 5].Value = AmountOfMoney += amount;
                    excelPackage.SaveAs("ClientInfo.xlsx",
                        ExcelMethodGroup.SetPassword("PasswordClient.xlsx", "password"));
                    return AmountOfMoney;
                }
            }
            return AmountOfMoney;
        }

        /// <summary>
        /// Deposit money into an account (database is XmlFile)
        /// </summary>
        /// <param name="amount"></param>
        /// <returns></returns>
        public double AddMoneyXml(double amount = 0)
        {
            XDocument xAccountDocument = XDocument.Load(XmlMethodGroup.ACCOUNTS_INFO_DOCUMENT);
            XElement xAccountElement = xAccountDocument.Element(XmlMethodGroup.ACCOUNTS_INFO_ELEM);

            foreach (var person in xAccountElement.Elements())
            {
                if (int.Parse(person.Attribute(XmlMethodGroup.ID).Value) == ID &&
                    amount == 0)
                {
                    person.Element(XmlMethodGroup.AMOUNT_OF_MONEY).Value =
                        (AmountOfMoney += InitializationHelper.DoubleInit("sum of money to add")).ToString();
                    xAccountDocument.Save(XmlMethodGroup.ACCOUNTS_INFO_DOCUMENT);
                    return AmountOfMoney;
                }
                else if (int.Parse(person.Attribute(XmlMethodGroup.ID).Value) == ID)
                {
                    person.Element(XmlMethodGroup.AMOUNT_OF_MONEY).Value =
                       (AmountOfMoney += amount).ToString();
                    xAccountDocument.Save(XmlMethodGroup.ACCOUNTS_INFO_DOCUMENT);
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
        public async Task<double> TakeMoneyAsync(double desiredAmount = 0)
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
                // return AmountOfMoney = await Task.Run(()=>
                //                         ExcelMethodGroup.WithdrawMoneyXLSXAsync(this, AmountOfMoney, tempDesAmount));

                return AmountOfMoney = XmlMethodGroup.WithdrawMoneyXml(this, AmountOfMoney, tempDesAmount);
            }
            //money check
            if (desiredAmount > AmountOfMoney)
            {
                MessageInformant.ErrorOutput($"Not enough money. You have {AmountOfMoney} BYN");
                return -1;
            }
            else
                return AmountOfMoney = await Task.Run(() =>
                      ExcelMethodGroup.WithdrawMoneyXLSXAsync(this, AmountOfMoney, desiredAmount));
        }
        /// <summary>
        /// Withdraw money from the account. Value is stored in xml file 
        /// </summary>
        /// <param name="desiredAmount">Desired amount of money to withdraw money</param>
        /// <returns></returns>
        public double TakeMoneyXml(double desiredAmount = 0)
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
                //Excel method
                //return AmountOfMoney = ExcelMethodGroup.WithdrawMoneyXLSX(this, AmountOfMoney, tempDesAmount);
                return AmountOfMoney = XmlMethodGroup.WithdrawMoneyXml(this, AmountOfMoney, tempDesAmount);
            }
            //money check
            if (desiredAmount > AmountOfMoney)
            {
                MessageInformant.ErrorOutput($"Not enough money. You have {AmountOfMoney} BYN");
                return -1;
            }
            else
                //Excel method
                //return AmountOfMoney = ExcelMethodGroup.WithdrawMoneyXLSX(this, AmountOfMoney, desiredAmount);
                return AmountOfMoney = XmlMethodGroup.WithdrawMoneyXml(this, AmountOfMoney, desiredAmount);
        }
        /// <summary>
        /// Check length login and password
        /// </summary>
        /// <param name="account"></param>
        /// <returns></returns>
        private static bool CheckLengthLoginAndPassword(Account account)
        {
            Type? type = typeof(Account);
            object[] attributes = type.GetCustomAttributes(false);
            foreach (Attribute attr in attributes)
            {
                if (attr is CheckLengthLoginPasswordAttribute attribute)
                {
                    if (attribute.Length > account.Login.Length && attribute.Length > account.Password.Length)
                        return true;
                    else if (attribute.Length < account.Login.Length)
                    {
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