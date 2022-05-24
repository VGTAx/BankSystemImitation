using InitHelperInformatMessage;
using Attributes;

namespace BankSystem
{
    /// <summary>
    /// сlass Account used to create a bank customer
    /// </summary>
    
    [CheckLengthLoginPassword(15)]  
    internal class Account : IAccount
    {        
        public event Func<string, double> doubleMethod;
        public event Func<string, string> stringMethod;
        public event Func<string,int> intMethod;

        public Account() { }
        public Account(Person person, string password, string login, double sum)
        {
            personObj = person;                      
            Login = login;
            Password = password;
            AmountOfMoney = sum;
        }
    
        public Person personObj { get; set; }
        public string Password { get; set; } 
        public string Login { get; set; }
        public string Name { get; set; }
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
            Console.WriteLine("Amount of money - {AmountOfMoney} BYN");
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
            doubleMethod += InitializationHelper.DoubleInit;
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

                personObj = new Person(agePerson,namePerson,surnamePerson);
                while(!Person.CheckAge(personObj))
                {                    
                    agePerson = intMethod.Invoke("age");
                    personObj = new Person(agePerson, namePerson, surnamePerson);
                }
            }
            while (CheckLength(new Account(personObj, Password, Login, AmountOfMoney))==false);

            MessageInformant.SuccessOutput("Account registered!");
            return new Account(personObj, Password, Login, AmountOfMoney);
        }
       
        /// <summary>
        /// Deposit money into an account
        /// </summary>
        /// <param name="amount"></param>
        /// <returns></returns>
        public double AddMoney(double amount = 0)
        {
            if (amount == 0)
            {
                return AmountOfMoney += doubleMethod.Invoke("sum of money to add");
            }
            else
                return AmountOfMoney += amount;
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
    }
}
