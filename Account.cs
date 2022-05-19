using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InitHelperInformatMessage;

namespace myBank
{
    /// <summary>
    /// сlass Account used to create a bank customer
    /// </summary>
    internal class Account : IAccount
    {
        public event Func<string, double> doubleMethod;
        public event Func<string, string> stringMethod;

        public Account() { }
        public Account(string name, string surName,string password,string login, double sum)
        {
            Name = name;
            Surname = surName;
            Login = login;
            Password = password;
            AmountOfMoney = sum;           
        }

        public string Password { get; set; }
        public string Login { get; set; }
        public string Name { get; set; }
        public string Surname { get; set; }
        public double AmountOfMoney { get; set; }
        public bool Authorization { get; set; }

        public bool LoginAcc()
        {
            
            string login = InitializationHelper.StringInIt("Login: ").ToUpper();
            string pass = InitializationHelper.StringInIt("Password: ").ToUpper();

            int count = 3;
            while (login != Login || pass != Password && count > 0)
            {
                MessageInformant.ErrorOutput($"Invalid login or password! {count} attemp left");                

                login = InitializationHelper.StringInIt("Login: ").ToUpper();                
                pass = InitializationHelper.StringInIt("Password: ").ToUpper();

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
            Console.WriteLine($"Name - { Name}\nSurname - { Surname}\nAmount of money - {AmountOfMoney} BYN");
            Console.ReadLine();

            return $"Name - {Name}\nSurname - {Surname}\nAmount of money - {AmountOfMoney} BYN";
        }
        /// <summary>
        /// Register account customer
        /// </summary>
        /// <param name="obj"></param>
        /// <returns></returns>
        public Account RegistrAcc(List<IAccount>obj)
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
            while (stringMethod?.Invoke("repeat pass") !=Password)
            {
                MessageInformant.ErrorOutput("Passwords do not match");
            } 
            
            Name = stringMethod.Invoke("name");
            Surname = stringMethod.Invoke("surname");
           MessageInformant.SuccessOutput("Account registered!");

            return new Account(Name, Surname, Password, Login, AmountOfMoney);
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
                
                while(temp==0 || temp>AmountOfMoney) 
                {
                    if(temp==0)//while try to withdraw 0 BYN
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
    }
}
