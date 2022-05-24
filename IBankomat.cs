using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankSystem
{
    internal interface IBankomat
    {
        string Adress { get; set; }
        double AmountOfMoneyATM { get; set; }
        int NumberATM { get; set; }

        public Bankomat CreateATM(List<IBankomat> bankomats);
        void Info();
        void LoadMoney();
        void WithdrawMoney(IAccount account);
    }
}
