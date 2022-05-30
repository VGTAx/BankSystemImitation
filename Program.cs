using OfficeOpenXml;

namespace BankSystem
{
    public class Program
    {
        private static void Main()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            Bank bank = new Bank();
            bank.ManageMain();
            Console.ReadLine();
        }
    }
}

