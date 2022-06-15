using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankSystem
{
    public interface IAccount
    {
        public int ID { get; set; }
        Person person { get; set; }
        string Login { get; set; }
        string Password { get; set; }
        double AmountOfMoney { get; set; }       
        bool Authorization { get; set; }    
        
        double AddMoney(double amount=0);
        async Task<double> TakeMoneyAsync(double amount = 0) 
        { 
            await Task.Run(() => null);
            return amount;
        }
        double TakeMoney(double amount = 0);
        public Account RegistrAcc(List<IAccount> obj);        
        string Info();
    }
}
