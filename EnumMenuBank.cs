namespace BankSystem
{
    /// <summary>
    /// Enum of methods for manage account of client
    /// </summary>
    public enum EnumManageAccountClient
    {
        AddMoney = 1,
        TakeMoney = 2,
        GetInfo,
        Back = 0
    }
    /// <summary>
    /// Enum of methods for manage bank
    /// </summary>
    public enum EnumManageBank
    {
        AddATM = 1,
        GetAllATM,
        LoadMoneyATM,
        Back = 0
    };
    /// <summary>
    /// Enum of methods for Main menu
    /// </summary>
    public enum EnumManageMain
    {
        Client = 1,
        Bank = 2,
        Exit = 0
    };
    /// <summary>
    /// Enum of methods to start working with a client account
    /// </summary>
    public enum EnumManageClient
    {
        SighUp = 1,
        SignIn,
        Back = 0
    }
    /// <summary>
    /// Enum properties class ATM
    /// </summary>
    public enum EnumATM
    {
        Number = 1,
        Adress = 2,
        MoneyATM = 3,
    }
    /// <summary>
    /// Enum properties class Acc (info about client)
    /// </summary>
    public enum EnumClient
    {
        ID = 1,
        Name = 2,        
        Surname = 3,        
        Age = 4,
        Money = 5,        
    }
    /// <summary>
    /// Enum properties class Acc (for login)
    /// </summary>
    public enum EnumAcc
    {
        Login = 2,
        Password = 3,
    }
}
