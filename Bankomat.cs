using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using InitHelperInformatMessage;

namespace myBank
{
    internal class Bankomat : IBankomat
    {
        public event Func<string, double> doubleMethod;
        public event Func<string, string> stringMethod;
        public Bankomat() 
        {
            doubleMethod += InitializationHelper.DoubleInit;
            stringMethod += InitializationHelper.StringInIt;
        }
        public Bankomat(string adress, double sum, int numberATM)
        {
            Adress=adress;
            AmountOfMoneyATM=sum;
            NumberATM=numberATM;
        }

        public string Adress { get; set; }
        public int NumberATM { get; set; }
        public double AmountOfMoneyATM { get; set; }
        /// <summary>
        /// Load money into ATM
        /// </summary>
        public void LoadMoney()
        {
            AmountOfMoneyATM += doubleMethod.Invoke("the amount of money to load into the ATM");           
        }
        /// <summary>
        /// Create ATM
        /// </summary>
        /// <param name="bankomats"></param>
        /// <returns></returns>
        public Bankomat CreateATM(List<IBankomat>bankomats)
        {
            Adress = stringMethod.Invoke("adress ATM");
            AmountOfMoneyATM = doubleMethod.Invoke("amount of money");
            NumberATM = (int)doubleMethod.Invoke("number ATM");

            foreach (var itLog in bankomats)
            {
                while (NumberATM == itLog.NumberATM)
                {
                    MessageInformant.ErrorOutput($"Number ATM \"{NumberATM}\" not available");
                    NumberATM = (int)doubleMethod.Invoke("number ATM");
                }
            }
            return new Bankomat(Adress, AmountOfMoneyATM, NumberATM);
        }
        /// <summary>
        /// Get information about ATM
        /// </summary>
        public void Info()
        {
            Console.WriteLine($"Adress: {Adress}");
            //Color for different amount of money
            //Red <3000, Yellow >3000&&<7000, Green>7000
            if(AmountOfMoneyATM < 3000)
            {
                Console.Write($"Amount of money in ATM: ");
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"{AmountOfMoneyATM} BYN");
                Console.ResetColor();
            }
            if(AmountOfMoneyATM > 3000 && AmountOfMoneyATM < 7000)
            {
                Console.Write($"Amount of money in ATM: ");
                Console.ForegroundColor = ConsoleColor.Yellow;
                Console.WriteLine($"{AmountOfMoneyATM} BYN");
                Console.ResetColor();
            }
            if (AmountOfMoneyATM > 7000)
            {
                Console.Write($"Amount of money in ATM: ");
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"{AmountOfMoneyATM} BYN");
                Console.ResetColor();
            }

            Console.Write($"Number ATM:");
            Console.ForegroundColor= ConsoleColor.Yellow;
            Console.WriteLine($" №{NumberATM}");
            Console.ResetColor();
        }
        /// <summary>
        /// Withdraw money from ATM
        /// </summary>
        /// <param name="account"></param>
        public void WithdrawMoney(IAccount account)
        {
            double temp = 0;
            bool check = true;
          
            temp = doubleMethod.Invoke("amount of money to withdraw");
            while(temp==0)
            {
                MessageInformant.ErrorOutput($"Can't withdraw 0 BYN");
                temp = doubleMethod.Invoke("sum of money to withdraw");
            }

            while (temp >= AmountOfMoneyATM || temp>account.AmountOfMoney)
            {
                if(!check)
                {
                    temp = doubleMethod.Invoke("amount of money to withdraw");
                }

                if (temp > AmountOfMoneyATM && AmountOfMoneyATM !=0)
                { 
                    MessageInformant.ErrorOutput($"There is not enough money in the ATM.Enter other amount of money");
                    check = false;
                    continue;
                }
                else if(AmountOfMoneyATM==0)
                {
                    MessageInformant.ErrorOutput($"There is not enough money in the ATM. Load money into ATM");
                    Thread.Sleep(2000);
                    check = false;
                    break;
                }
                else if (account.AmountOfMoney < temp)
                {
                    MessageInformant.ErrorOutput($"You have not enough money ({account.AmountOfMoney} BYN)");
                    check = false;
                    continue;
                }
                else
                {
                    check = true;                    
                    break;
                }
            }
            if(check)
            {
                MessageInformant.SuccessOutput($"Money withdrawn {temp} BYN");
                account.TakeMoney(temp);
                AmountOfMoneyATM -= temp;
            }
        }
    }
}