using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BankSystem
{
    internal interface IBank
    {      
        private string Name 
        { 
            get { return Name; }            
        }
        private List<IAccount> Accounts { get { return new List<IAccount>(); }  }
        private List<IATM> Bankomats { get { return new List<IATM>(); } }

        private void SighUp()  { }
        private void SignIn() { }
        private void AddMoney() { }
        private void TakeMoney() { }
        private void GetInfoClient() { }

        private void LoadMoneyATM() { }        
        private void AddATM() { }
        private void GetAllATM()  { }

        private void ManageAccountClient() { }
        private void ManageClient() { }
        private void ManageBank() { }
        private void ManageMain() { }
    }
}
