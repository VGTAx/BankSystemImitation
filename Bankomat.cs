using InitHelperInformatMessage;
using OfficeOpenXml;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.Remoting;
using System.Diagnostics;

namespace BankSystem
{
    public sealed class Bankomat : IBankomat
    {
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
            FileInfo fileInfo = new FileInfo("ATMInfo.xlsx");
            ExcelPackage packageATM = new ExcelPackage(fileInfo, ExcelMethodGroup.GetPassword("PasswordATM.xlsx", "password"));
            ExcelWorksheet worksheetATM = packageATM.Workbook.Worksheets["ATM Info"];

            int rowNumber = 0;
            for (int i = worksheetATM.Dimension.Start.Row+1; i <= worksheetATM.Dimension.End.Row; i++)
            {
                for (int j = worksheetATM.Dimension.Start.Column; j <= worksheetATM.Dimension.Start.Column; j++)
                {
                    int tempNumberATM = int.Parse(worksheetATM.Cells[i, j].Value.ToString());
                    if (tempNumberATM == NumberATM)
                    {
                        rowNumber = i;
                    }                    
                }
            }
            worksheetATM.Cells[rowNumber, 3].Value = AmountOfMoneyATM;
            packageATM.SaveAs("ATMInfo.xlsx", 
                ExcelMethodGroup.SetPassword("PasswordATM.xlsx", "password"));

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
            NumberATM = new Random().Next(1, 999);            
            ExcelMethodGroup.WorksheetAtmXLSXAsync(new Bankomat(Adress, AmountOfMoneyATM, NumberATM));
            MessageInformant.SuccessOutput("ATM Added!");
            Console.ReadLine();
            return new Bankomat(Adress, AmountOfMoneyATM, NumberATM);
        }
        /// <summary>
        /// Get information about ATM
        /// </summary>
        public void Info()
        {
            Console.WriteLine($"Adress: {Adress}");

            ///Color for different amount of money
            ///Red <3000, Yellow >3000&&<7000, Green>7000
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
            Console.WriteLine();
        }
        /// <summary>
        /// Withdraw money from ATM
        /// </summary>
        /// <param name="account"></param>
        public async void WithdrawMoney(IAccount account)
        {
            
            double tempDesAmount = 0;
            bool check = true;

            //tempDesAmount = InitializationHelper.DoubleInit("amount of money to withdraw");
            tempDesAmount = 1200;
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
                    MessageInformant.ErrorOutput($"There is not enough money in the ATM.Enter other amount of money" +
                        $". Available amount of money {AmountOfMoneyATM} BYN!");
                    check = false;
                    continue;
                }
                else if (AmountOfMoneyATM == 0)
                {
                    MessageInformant.ErrorOutput($"There is not enough money in the ATM. Load money into ATM!");
                    Thread.Sleep(2000);
                    check = false;
                    break;
                }
                else if (account.AmountOfMoney < tempDesAmount)
                {
                    MessageInformant.ErrorOutput($"You have not enough money ({account.AmountOfMoney} BYN!)");
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
                AmountOfMoneyATM = await ExcelMethodGroup.WithdrawMoneyAtmXLSXAsync(this, AmountOfMoneyATM, tempDesAmount);
                account.TakeMoneyAsync(tempDesAmount);  
            }
        }
    }
}