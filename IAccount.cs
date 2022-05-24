using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankSystem
{
    internal interface IAccount
    {
        string Name { get; set; }
        string Surname { get; set; }

        string Login { get; set; }
        string Password { get; set; }
        double AmountOfMoney { get; set; }

        bool Authorization { get; set; }
        double AddMoney(double amount=0);
        double TakeMoney(double amount=0);
        public Account RegistrAcc(List<IAccount> obj);
        bool LoginAcc();
        string Info();
    }
}
