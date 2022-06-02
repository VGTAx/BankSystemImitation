using InitHelperInformatMessage;

namespace BankSystem
{
    internal sealed class Bank : IBank
    {
        private delegate void Manage();

        private Dictionary<int, Manage> DictManageBank;
        private Dictionary<int, Manage> DictManageClient;
        private Dictionary<int, Manage> DictManageMain;
        private Dictionary<int, Manage> DictManageAccountClient;
        /// <summary>
        /// Enum of methods for manage account of client
        /// </summary>
        private enum EnumManageAccountClient
        {
            AddMoney = 1,
            TakeMoney = 2,
            GetInfo,
            Back = 0
        }
        /// <summary>
        /// Enum of methods for manage bank
        /// </summary>
        private enum EnumManageBank
        {
            AddATM = 1,
            GetAllATM,
            LoadMoneyATM,
            Back = 0
        };
        /// <summary>
        /// Enum of methods for Main menu
        /// </summary>
        private enum EnumManageMain
        {
            Client = 1,
            Bank = 2,
            Exit = 0
        };
        /// <summary>
        /// Enum of methods to start working with a client account
        /// </summary>
        private enum EnumManageClient
        {
            SighUp = 1,
            SignIn,
            Back = 0
        }

        public Bank()
        {
            Name = String.Empty;
            Accounts = new List<IAccount>();
            Bankomats = new List<IBankomat>();
        }
        public Bank(string name, List<IAccount> acc, List<IBankomat> atm)
        {
            Name = name;
            Accounts = acc;
            Bankomats = atm;
        }

        private string Name { get; set; }
        private List<IAccount> Accounts { get; set; }
        private List<IBankomat> Bankomats { get; set; }
        /// <summary>
        /// Sign Up account of client
        /// </summary>
        private void SighUp()
        {
            Console.Clear();
            Account account = new Account();
            account.RegistrAcc(Accounts);
            Accounts.Add(account);
        }
        /// <summary>
        /// Login method
        /// </summary>
        private void SignIn()
        {
            Console.Clear();
            int attemptCount = 3;//attemp count
            bool check = false;
            Accounts = ExcelMethodGroup.LoadListAccXLSX();
            //checking for accounts
            if (Accounts == null | Accounts.Count == 0)
            {
                MessageInformant.ErrorOutput("There is no Account. Sigh Up account first!");
                check = true;
            }
            while (!check)
            {
                string login = InitializationHelper.StringInIt("Login");
                string pass = InitializationHelper.StringInIt("Password");
                //authorization client
                var client = from account in Accounts
                             where account.Login == login
                             where account.Password == pass
                             select account;

                foreach (var db in client)
                {
                    check = true;
                    db.Authorization = true;
                    MessageInformant.SuccessOutput("Login succeeded");
                    ManageAccountClient();
                    db.Authorization = false;
                }
                //decrease in the number of login attempt, if invalid login or password
                if (!check)
                {
                    MessageInformant.ErrorOutput($"Invalid login or password! {attemptCount} attemp left");
                    attemptCount--;
                }
                //Error, if attempts are over
                if (attemptCount == 0)
                {
                    MessageInformant.ErrorOutput("attempt limit reached");
                    check = true;
                }
            }
        }
        /// <summary>
        /// Put money into account client
        /// </summary>
        private void AddMoney()
        {

            var client = from p in Accounts where p.Authorization == true select p;

            foreach (var item in client)
            {
                item.AddMoney();
                MessageInformant.SuccessOutput("Money added to account!");
            }
        }
        /// <summary>
        /// Withdraw money by client
        /// </summary>
        private void TakeMoney()
        {
            //Finding an "Authorization Key"
            var clientTakeMoney = from p in Accounts where p.Authorization == true select p;

            foreach (var accAUTH in clientTakeMoney)
            {
                //checking if the client has money
                if (accAUTH.AmountOfMoney != 0)
                {
                    while (true)
                    {
                        //select where withdraw money
                        Console.Write("Select where do you wish to withdraw money:\n1.Bank\n2.ATM\n");
                        string enter = Console.ReadLine();
                        switch (enter)
                        {
                            case "1":
                                accAUTH.TakeMoney();
                                MessageInformant.SuccessOutput($"Money withdrawn");
                                Console.ReadLine();
                                break;
                            case "2":
                                while (true)
                                {   //checking if the bank has ATM
                                    Bankomats = ExcelMethodGroup.LoadListAtmXLSX();
                                    if (Bankomats.Count == 0)
                                    {
                                        MessageInformant.ErrorOutput("There is no ATM. Add ATM first!");
                                        Console.ReadLine();
                                        break;
                                    }

                                    while (true)
                                    {   //get list ATM

                                        foreach (var atm in Bankomats)
                                        {
                                            atm.Info();
                                        }
                                        //select ATM
                                        Console.Write("\nSelect ATM to withdraw money (");
                                        int tempATM = (int)InitializationHelper.DoubleInit("№ATM)");
                                        //request to find an ATM
                                        var atmSelect = from p in Bankomats where p.NumberATM == tempATM select p;

                                        foreach (var authATM in atmSelect)
                                        {
                                            authATM.WithdrawMoney(accAUTH);
                                            MessageInformant.SuccessOutput($"Money withdrawn");
                                            Console.ReadLine();
                                        }
                                        if (atmSelect.Any())
                                        {
                                            break;
                                        }

                                        MessageInformant.ErrorOutput("Invalid ATM Number!");
                                        Console.ReadLine();
                                        Console.Clear();
                                    }
                                    break;
                                }
                                break;
                        }
                        break;
                    }
                }
                else
                {
                    //while client has not money
                    MessageInformant.ErrorOutput($"You have {accAUTH.AmountOfMoney} BYN. Top up your account!");
                    Console.ReadLine();
                }
            }
        }
        /// <summary>
        /// Get information about the client
        /// </summary>
        private void GetInfoClient()
        {
            var client = from p in Accounts
                         where p.Authorization == true
                         select p;

            foreach (var item in client)
            {
                item.Info();
            }
        }

        /// <summary>
        /// Add ATM of Bank
        /// </summary>
        private void AddATM()
        {
            Bankomat atm = new Bankomat();
            atm.CreateATM(Bankomats);
            Bankomats.Add(atm);
            MessageInformant.SuccessOutput("ATM Add!");
            Console.Clear();
        }
        /// <summary>
        /// Get list ATM
        /// </summary>
        private void GetAllATM()
        {
            Console.Clear();
            Bankomats = ExcelMethodGroup.LoadListAtmXLSX();
            if (Bankomats.Count != 0)
            {
                var listATM = from p in Bankomats select p;

                foreach (var item in listATM)
                {
                    item.Info();
                }
            }
            else
            {
                MessageInformant.ErrorOutput("There is no ATM. Add ATM first!");
                Console.ReadLine();
                Console.Clear();
            }
        }
        /// <summary>
        /// Load money into ATM
        /// </summary>
        private void LoadMoneyATM()
        {
            bool check = false;
            Bankomats = ExcelMethodGroup.LoadListAtmXLSX();
            while (Bankomats.Count != 0)
            {
                GetAllATM();
                Console.WriteLine();
                int ATM = InitializationHelper.IntInit("number ATM!");

                var selectATM = from p in Bankomats where p.NumberATM == ATM select p;//search entered ATM

                foreach (var atm in selectATM)
                {
                    atm.LoadMoney();
                    check = true;
                }

                if (check)
                {
                    MessageInformant.SuccessOutput("Money load into the ATM!");
                    Console.ReadLine();
                    Console.Clear();
                    break;
                }
                else
                {
                    MessageInformant.ErrorOutput("Invalid number ATM!");
                    Console.ReadLine();
                }
            }
            if (Bankomats.Count == 0)
            {
                MessageInformant.ErrorOutput("There is no ATM. Add ATM first!");
                Console.ReadLine();
                Console.Clear();
            }
        }
        /// <summary>
        /// Manage Account client
        /// </summary>
        private void ManageAccountClient()
        {
            Console.Clear();
            //create a dictionary to store methods
            DictManageAccountClient = new Dictionary<int, Manage>
            {
                {1, new Manage(AddMoney) },
                {2, new Manage(TakeMoney) },
                {3, new Manage(GetInfoClient) }
            };
            //created array that enum available method
            EnumManageAccountClient enumManageAccountClient = new EnumManageAccountClient();
            Array EnumManageAccountClient = Enum.GetValues(enumManageAccountClient.GetType());

            while (true)
            {
                Console.Clear();
                Console.ForegroundColor = ConsoleColor.DarkCyan;
                Console.WriteLine("1. Добавление денег на счет" +
                    "\n2. Снятие денег со счета\n3. Получение общей информации");
                Console.ResetColor();
                Console.WriteLine("Select function(Enter number):\n");
                //enum available method
                foreach (int item in EnumManageAccountClient)
                {
                    Console.WriteLine($"{item}.{EnumManageAccountClient.GetValue(item)}");
                }

                bool result = int.TryParse(Console.ReadLine(), out int select);

                if (result && select == 0)
                {
                    Console.Clear();
                    break;
                }
                //passing selected method to delegate
                if (result && EnumManageAccountClient.Length > select)
                {
                    DictManageAccountClient[select]();
                }
                else
                {
                    MessageInformant.ErrorOutput("Invalid select");
                    Thread.Sleep(400);
                }
            }
        }

        /// <summary>
        /// Manage Bank
        /// </summary>
        private void ManageBank()
        {
            Console.Clear();
            //create a dictionary to store methods
            DictManageBank = new Dictionary<int, Manage>
            {
                {1, new Manage(AddATM) },
                {2, new Manage(GetAllATM) },
                {3, new Manage(LoadMoneyATM) }
            };
            Console.Clear();
            //created array that enum available method
            EnumManageBank manageBank = new EnumManageBank();
            Array EnumManageBank = Enum.GetValues(manageBank.GetType());

            while (true)
            {
                Console.ForegroundColor = ConsoleColor.DarkCyan;
                Console.WriteLine("1. Добавить банкомат" +
                    "\n2. Список банкоматов\n3. Загрузить деньги в банкомат");
                Console.ResetColor();
                Console.WriteLine("Select function(Enter number):\n");
                //enum available method
                foreach (int item in EnumManageBank)
                {
                    Console.WriteLine($"{item}.{EnumManageBank.GetValue(item)}");
                }
                bool result = int.TryParse(Console.ReadLine(), out int select);

                if (result && select == 0)
                {
                    Console.Clear();
                    break;
                }
                //passing selected method to delegate
                if (result && EnumManageBank.Length > select)
                {
                    DictManageBank[select]();
                }
                else
                {
                    MessageInformant.ErrorOutput("Invalid select");
                    Thread.Sleep(400);
                }
            }
        }
        /// <summary>
        /// Manage client's personal account
        /// </summary>
        private void ManageClient()
        {
            Console.Clear();
            //create a dictionary to store methods
            DictManageClient = new Dictionary<int, Manage>
            {
                {1, new Manage(SighUp) },
                {2, new Manage(SignIn) },
            };
            //created array that enum available method
            EnumManageClient manageClient = new EnumManageClient();
            Array EnumManageClient = Enum.GetValues(manageClient.GetType());


            while (true)
            {
                Console.ForegroundColor = ConsoleColor.DarkCyan;
                Console.WriteLine("1. Регистрация\n2. Вход в личный кабинет ");
                Console.ResetColor();
                Console.WriteLine("Select function(Enter number):\n");
                //enum available method
                foreach (int item in EnumManageClient)
                {
                    Console.WriteLine($"{item}.{EnumManageClient.GetValue(item)}");
                }
                bool result = int.TryParse(Console.ReadLine(), out int select);

                if (result && select == 0)
                {
                    Console.Clear();
                    break;
                }
                //passing selected method to delegate
                if (result && EnumManageClient.Length > select)
                {
                    DictManageClient[select]();
                }
                else
                {
                    MessageInformant.ErrorOutput("Invalid select");
                    Thread.Sleep(400);
                }
            }
        }
        /// <summary>
        /// Main menu
        /// </summary>
        public void ManageMain()
        {
            Console.Clear();
            //create a dictionary to store methods
            DictManageMain = new Dictionary<int, Manage>
            {
                {1,new Manage(ManageClient)},
                {2,new Manage(ManageBank)}
            };
            //created array that enum available method
            EnumManageMain manageMain = new EnumManageMain();
            Array EnumManageMain = Enum.GetValues(manageMain.GetType());

            while (true)
            {
                Console.Clear();
                Console.ForegroundColor = ConsoleColor.DarkCyan;
                Console.WriteLine("1. Личный кабинет пользователя" +
                    "\n2. Управление банком");
                Console.ResetColor();
                Console.WriteLine("Select user(Enter number):\n");
                //enum available method
                foreach (int item in EnumManageMain)
                {
                    Console.WriteLine($"{item}.{EnumManageMain.GetValue(item)}");
                }
                bool result = int.TryParse(Console.ReadLine(), out int select);

                if (result && select == 0)
                {
                    Console.Clear();
                    string temp = InitializationHelper.StringInIt("\"Y\" or \"y\" to exit or " +
                        "press any button to continue");
                    if (temp == "Y")
                        break;
                    continue;
                }
                //passing selected method to delegate
                if (result && EnumManageMain.Length > select)
                {
                    DictManageMain[select]();
                }
                else
                {
                    MessageInformant.ErrorOutput("Invalid select");
                    Console.ReadLine();
                }
            }
        }
    }
}