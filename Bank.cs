using InitHelperInformatMessage;

namespace BankSystem
{
    public sealed class Bank : IBank
    {
        private delegate void Manage();

        private Dictionary<EnumManageBank, Manage> DictManageBank;
        private Dictionary<EnumManageClient, Manage> DictManageClient;
        private Dictionary<EnumManageMain, Manage> DictManageMain;
        private Dictionary<EnumManageAccountClient, Manage> DictManageAccountClient;

        public Bank()
        {
            Name = String.Empty;
            Accounts = new List<IAccount>();
            ATM = new List<IATM>();
        }
        public Bank(string name, List<IAccount> acc, List<IATM> atm)
        {
            Name = name;
            Accounts = acc;
            ATM = atm;
        }

        private string Name { get; set; }
        private List<IAccount> Accounts { get; set; }
        private List<IATM> ATM { get; set; }
        /// <summary>
        /// Sign Up account of client
        /// </summary>
        private void SighUp()
        {
            Console.Clear();
            Account account = new Account();
            account.RegistrAcc(Accounts);
            Accounts.Add(new Account());
        }
        /// <summary>
        /// Login method
        /// </summary>
        private void SignIn()
        {
            Console.Clear();
            int attemptCount = 3;//attemp count
            bool check = false;
            //Accounts = ExcelMethodGroup.LoadListAccXLSX();
            Accounts = XmlMethodGroup.LoadXmlListAccount();

            //checking for accounts
            if (Accounts == null | Accounts.Count == 0)
            {
                MessageInformant.ErrorOutput("There is no Account. Sigh Up account first!");
                Console.ReadLine();
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
                    MessageInformant.SuccessOutput("Login succeeded!");
                    Console.ReadLine();
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
                //item.AddMoney();
                item.AddMoneyXml();
                MessageInformant.SuccessOutput("Money added to account!");
            }
        }
        /// <summary>
        /// Withdraw money by client
        /// </summary>
        private void WithdrawMoney()
        {
            //Finding an "Authorization Key"
            var clientWithdrawMoney = from p in Accounts
                                      where p.Authorization == true && p.AmountOfMoney != 0
                                      select p;
            if (clientWithdrawMoney.Any() == false)
            {
                //while client has not money
                MessageInformant.ErrorOutput($"You have 0 BYN. Top up your account!");
                Console.ReadLine();
            }
            foreach (var accAUTH in clientWithdrawMoney)
            {
                //select where withdraw money
                Console.Write("Select where do you wish to withdraw money:\n1.Bank\n2.ATM\n");
                string enter = Console.ReadLine();
                switch (enter)
                {
                    case "1":
                        accAUTH.TakeMoneyXml();
                        MessageInformant.SuccessOutput($"Money withdrawn!");
                        Console.ReadLine();
                        break;
                    case "2":
                        ///checking if the bank has ATM (Excel file database)
                        //ATM = ExcelMethodGroup.LoadListAtmXLSX();

                        ///checking if the bank has ATM (Xml file database)
                        ATM = XmlMethodGroup.LoadXmlListAtm();
                        if (ATM.Count == 0)
                        {
                            MessageInformant.ErrorOutput("There is no ATM. Add ATM first!");
                            Console.ReadLine();
                            break;
                        }
                        while (ATM.Count != 0)
                        {
                            ///get list ATM
                            foreach (var atm in ATM)
                            {
                                atm.Info();
                            }
                            ///select ATM
                            Console.Write("\nSelect ATM to withdraw money (");
                            int tempATM = (int)InitializationHelper.DoubleInit("№ATM) or 0 to exit");
                            ///request to find an ATM
                            var atmSelect = from p in ATM where p.NumberATM == tempATM select p;

                            foreach (var authATM in atmSelect)
                            {
                                authATM.WithdrawMoney(accAUTH);
                                MessageInformant.SuccessOutput($"Money withdrawn!");
                                Console.ReadLine();
                            }
                            if (atmSelect.Any() || tempATM == 0)
                            {
                                break;
                            }

                            MessageInformant.ErrorOutput("Invalid ATM Number!");
                            Console.ReadLine();
                            Console.Clear();
                        }
                        break;
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
            ATM atm = new ATM();
            atm.CreateATM(ATM);
            ATM.Add(atm);
            MessageInformant.SuccessOutput("ATM Add!");
            Console.Clear();
        }
        /// <summary>
        /// Get list ATM
        /// </summary>
        private void GetAllATM()
        {
            Console.Clear();
            //ATM = ExcelMethodGroup.LoadListAtmXLSX();
            ATM = XmlMethodGroup.LoadXmlListAtm();
            if (ATM.Count != 0)
            {
                var listATM = from p in ATM select p;

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
            //ATM = ExcelMethodGroup.LoadListAtmXLSX();
            ATM = XmlMethodGroup.LoadXmlListAtm();
            while (ATM.Count != 0)
            {
                GetAllATM();
                Console.WriteLine();
                int ATM = InitializationHelper.IntInit("number ATM");

                var selectATM = from p in this.ATM where p.NumberATM == ATM select p;//search entered ATM

                foreach (var atm in selectATM)
                {
                    atm.LoadMoneyXml();
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
            if (ATM.Count == 0)
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
            DictManageAccountClient = new Dictionary<EnumManageAccountClient, Manage>
            {
                {EnumManageAccountClient.AddMoney, new Manage(AddMoney) },
                {EnumManageAccountClient.TakeMoney, new Manage(WithdrawMoney) },
                {EnumManageAccountClient.GetInfo, new Manage(GetInfoClient) }
            };

            while (true)
            {
                Console.Clear();
                Console.ForegroundColor = ConsoleColor.DarkCyan;
                Console.WriteLine("1. Добавление денег на счет" +
                    "\n2. Снятие денег со счета\n3. Получение общей информации");
                Console.ResetColor();
                Console.WriteLine("Select function(Enter number):\n");
                //enum available method
                foreach (var item in Enum.GetValues(typeof(EnumManageAccountClient)))
                {
                    Console.WriteLine($"{(int)item}.{(EnumManageAccountClient)item}");
                }

                bool result = Enum.TryParse(Console.ReadLine(), out EnumManageAccountClient select);

                try
                {
                    if (result && select == 0)
                    {
                        Console.Clear();
                        break;
                    };
                    DictManageAccountClient[select]();
                }
                catch (Exception)
                {
                    MessageInformant.ErrorOutput("Invalid select");
                    Console.ReadLine();
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
            DictManageBank = new Dictionary<EnumManageBank, Manage>
            {
                {EnumManageBank.AddATM, new Manage(AddATM) },
                {EnumManageBank.GetAllATM, new Manage(GetAllATM) },
                {EnumManageBank.LoadMoneyATM, new Manage(LoadMoneyATM) }
            };
            Console.Clear();

            while (true)
            {
                Console.ForegroundColor = ConsoleColor.DarkCyan;
                Console.WriteLine("1. Добавить банкомат" +
                    "\n2. Список банкоматов\n3. Загрузить деньги в банкомат");
                Console.ResetColor();
                Console.WriteLine("Select function(Enter number):\n");
                //enum available method
                foreach (int item in Enum.GetValues(typeof(EnumManageBank)))
                {
                    Console.WriteLine($"{item}.{(EnumManageBank)item}");
                }
                bool result = Enum.TryParse(Console.ReadLine(), out EnumManageBank select);

                try
                {
                    if (result && select == 0)
                    {
                        Console.Clear();
                        break;
                    }
                    DictManageBank[select]();
                }
                catch (Exception)
                {
                    MessageInformant.ErrorOutput("Invalid select");
                    Console.ReadLine();
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
            DictManageClient = new Dictionary<EnumManageClient, Manage>
            {
                {EnumManageClient.SighUp, new Manage(SighUp) },
                {EnumManageClient.SignIn, new Manage(SignIn) },
            };

            while (true)
            {
                Console.Clear();
                Console.ForegroundColor = ConsoleColor.DarkCyan;
                Console.WriteLine("1. Регистрация\n2. Вход в личный кабинет ");
                Console.ResetColor();
                Console.WriteLine("Select function(Enter number):\n");
                //enum available method
                foreach (var item in Enum.GetValues(typeof(EnumManageClient)))
                {
                    Console.WriteLine($"{(int)item}.{(EnumManageClient)item}");
                }
                bool result = Enum.TryParse(Console.ReadLine(), out EnumManageClient select);

                try
                {
                    if (result && select == 0)
                    {
                        Console.Clear();
                        break;
                    }
                    DictManageClient[select]();
                }
                catch (Exception)
                {
                    MessageInformant.ErrorOutput("Invalid select");
                    Console.ReadLine();
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
            DictManageMain = new Dictionary<EnumManageMain, Manage>
            {
                {EnumManageMain.Client,new Manage(ManageClient)},
                {EnumManageMain.Bank,new Manage(ManageBank)}
            };

            while (true)
            {
                Console.Clear();
                Console.ForegroundColor = ConsoleColor.DarkCyan;
                Console.WriteLine("1. Личный кабинет пользователя" +
                    "\n2. Управление банком");
                Console.ResetColor();
                Console.WriteLine("Select user(Enter number):\n");
                //enum available method
                foreach (var item in Enum.GetValues(typeof(EnumManageMain)))
                {
                    Console.WriteLine($"{(int)item}.{(EnumManageMain)item}");
                }
                bool result = Enum.TryParse(Console.ReadLine(), out EnumManageMain select);

                //passing selected method to delegate
                try
                {
                    if (result && select == 0)
                    {
                        Console.Clear();
                        string temp = InitializationHelper.StringInIt("\"Y\" or \"y\" to exit or " +
                            "press any button to continue");
                        if (temp == "Y")
                            break;
                        continue;
                    }
                    DictManageMain[select]();
                }
                catch
                {
                    MessageInformant.ErrorOutput("Invalid select");
                    Console.ReadLine();
                }
            }
        }
    }
}