namespace BankSystem
{
    public interface IBankomat
    {
        string Adress { get; set; }
        double AmountOfMoneyATM { get; set; }
        int NumberATM { get; set; }

        public Bankomat CreateATM(List<IBankomat> bankomats);
        void Info();
        void LoadMoney();
        public void WithdrawMoney(IAccount account);
    }
}
