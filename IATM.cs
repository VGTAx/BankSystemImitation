namespace BankSystem
{
    public interface IATM
    {
        string Adress { get; set; }
        double AmountOfMoneyATM { get; set; }
        int NumberATM { get; set; }

        public ATM CreateATM(List<IATM> atm);
        void Info();
        void LoadMoney();
        void LoadMoneyXml();
        public void WithdrawMoney(IAccount account);
    }
}
