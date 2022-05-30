using InitHelperInformatMessage;

namespace BankSystem
{
    internal class Bankomat : IBankomat
    {
        // public event Func<string, double> doubleMethod;
        //public event Func<string, string> stringMethod;
        public Bankomat() { }
        public Bankomat(string adress, double sum, int numberATM)
        {
            Adress = adress;
            AmountOfMoneyATM = sum;
            NumberATM = numberATM;
        }

        public string Adress { get; set; }
        public int NumberATM { get; set; }
        public double AmountOfMoneyATM { get; set; }
        /// <summary>
        /// Load money into ATM
        /// </summary>
        public void LoadMoney()
        {
            AmountOfMoneyATM += InitializationHelper.DoubleInit("the amount of money to load into the ATM");
        }
        /// <summary>
        /// Create ATM
        /// </summary>
        /// <param name="bankomats"></param>
        /// <returns></returns>
        public Bankomat CreateATM(List<IBankomat> bankomats)
        {
            Adress = InitializationHelper.StringInIt("adress ATM");
            AmountOfMoneyATM = InitializationHelper.DoubleInit("amount of money");
            NumberATM = (int)InitializationHelper.DoubleInit("number ATM");

            foreach (var itLog in bankomats)
            {
                while (NumberATM == itLog.NumberATM)
                {
                    MessageInformant.ErrorOutput($"Number ATM \"{NumberATM}\" not available");
                    NumberATM = (int)InitializationHelper.DoubleInit("number ATM");
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
            if (AmountOfMoneyATM < 3000)
            {
                Console.Write($"Amount of money in ATM: ");
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine($"{AmountOfMoneyATM} BYN");
                Console.ResetColor();
            }
            if (AmountOfMoneyATM > 3000 && AmountOfMoneyATM < 7000)
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
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($" №{NumberATM}");
            Console.ResetColor();
        }
        /// <summary>
        /// Withdraw money from ATM
        /// </summary>
        /// <param name="account"></param>
        public void WithdrawMoney(IAccount account)
        {
            double tempDesAmount = 0;
            bool check = true;

            tempDesAmount = InitializationHelper.DoubleInit("amount of money to withdraw");
            while (tempDesAmount == 0)
            {
                MessageInformant.ErrorOutput($"Can't withdraw 0 BYN");
                tempDesAmount = InitializationHelper.DoubleInit("sum of money to withdraw");
            }

            while (tempDesAmount >= AmountOfMoneyATM || tempDesAmount > account.AmountOfMoney)
            {
                if (!check)
                {
                    tempDesAmount = InitializationHelper.DoubleInit("amount of money to withdraw");
                }

                if (tempDesAmount > AmountOfMoneyATM && AmountOfMoneyATM != 0)
                {
                    MessageInformant.ErrorOutput($"There is not enough money in the ATM.Enter other amount of money");
                    check = false;
                    continue;
                }
                else if (AmountOfMoneyATM == 0)
                {
                    MessageInformant.ErrorOutput($"There is not enough money in the ATM. Load money into ATM");
                    Thread.Sleep(2000);
                    check = false;
                    break;
                }
                else if (account.AmountOfMoney < tempDesAmount)
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
            if (check)
            {
                AmountOfMoneyATM = ExcelMethodGroup.WithDrawMoney(account, AmountOfMoneyATM, tempDesAmount);
                MessageInformant.SuccessOutput($"Money withdrawn {tempDesAmount} BYN");
                account.TakeMoney(tempDesAmount);  
            }
        }
    }
}