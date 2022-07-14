using InitHelperInformatMessage;
using System.Xml.Linq;

namespace BankSystem
{
    internal class XmlMethodGroup
    {
        /// <summary>
        /// Name document(file) where stored accounts login info (login and password)
        /// </summary>
        public const string ACCOUNTS_LOGIN_INFO_DOCUMENT = "AccountLoginInfo.xml";
        /// <summary>
        /// Name root elem of "ACCOUNTS_LOGIN_INFO_DOCUMENT"  where stored list of Accounts (Login info)
        /// </summary>
        public const string ACCOUNTS_LOGIN_INFO_ELEM = "AccountInfo";

        /// <summary>
        /// Name document(file) where stored information about Person (Name,Surname,Age, Amount of money and so on)
        /// </summary>
        public const string ACCOUNTS_INFO_DOCUMENT = "Accounts.xml";
        /// <summary>
        /// Name root elem of "ACCOUNTS_INFO_DOCUMENT" where stored list of Person
        /// </summary>
        public const string ACCOUNTS_INFO_ELEM = "Accounts";
        /// <summary>
        /// Name document(file) where stored ATM info (Number, Adress, Amount of Money)
        /// </summary>
        public const string ATM_INFO_DOCUMENT = "AtmInfo.xml";
        /// <summary>
        /// Name root elem of "ATM_INFO_DOCUMENT"  where stored list of ATM
        /// </summary>
        public const string ATM_INFO_ELEM = "AtmInfo";

        public const string NUMBER_ATM = "NumberAtm";
        public const string ADRESS = "Adress";
        public const string ID = "ID";
        public const string NAME = "Name";
        public const string SURNAME = "Surname";
        public const string AGE = "Age";
        public const string AMOUNT_OF_MONEY = "AmountOfMoney";
        public const string PASSWORD = "Password";
        public const string LOGIN = "Login";

        /// <summary>
        /// Checking if a value is available (ID or Number Atm for example)
        /// </summary>
        /// <param name="value"> Checked value (ID or Number atm for example)</param>
        /// <param name="xmlDocName"> Name document(file) in which the availability of the value is checked and will be saved</param>
        /// <param name="xmlElem"> Name root elem of "xmlDocName" document </param>
        /// <param name="xmlAttribute">Name checked attribute (ID or Number for example)</param>
        /// <returns></returns>
        public static bool CheckValueAvailableXml(int value, string xmlDocName, string xmlElem, string xmlAttribute)
        {
            try
            {
                XDocument xDocument = XDocument.Load(xmlDocName);
                XElement xElement = xDocument.Element(xmlElem);

                foreach (var item in xElement.Elements())
                {
                    if (int.Parse(item.Attribute(xmlAttribute).Value) == value)
                        return false;
                }
                return true;
            }
            catch (Exception)
            {
                return true;
            }
        }
        /// <summary>
        /// Checking if an Account name (Login) is available 
        /// </summary>
        /// <param name="login">Checked value</param>
        /// <returns></returns>
        public static bool CheckAccNameAvailableXml(string login)
        {
            try
            {
                XDocument xAccountLoginInfoDocument = XDocument.Load(ACCOUNTS_LOGIN_INFO_DOCUMENT);
                XElement accounts = xAccountLoginInfoDocument.Element(ACCOUNTS_LOGIN_INFO_ELEM);

                foreach (var xLogin in accounts.Elements())
                {
                    if (xLogin.Element(LOGIN).Value.ToString() == login)
                    {
                        MessageInformant.ErrorOutput($"Login \"{login}\" not available");
                        return false;
                    }
                    else
                        continue;
                }
                return true;
            }
            catch (Exception)
            {
                return true;
            }
        }
        /// <summary>
        /// Opening or creating xml file with Accounts information
        /// </summary>
        /// <param name="account">Object for adding in xml file</param>
        public static void OpenOrCreateXmlAccountFile(IAccount account)
        {
            XDocument xAccountDocument = new XDocument();
            XDocument xAccountLoginInfoDocument = new XDocument();

            XElement xAccLoginInfoRootElem = new XElement(ACCOUNTS_LOGIN_INFO_ELEM);
            XElement xAccRootElem = new XElement(ACCOUNTS_INFO_ELEM);

            XElement xAccPerson = new XElement("Person",
                                                new XAttribute(ID, account.ID),
                                                new XElement(NAME, account.person.Name),
                                                new XElement(SURNAME, account.person.SurName),
                                                new XElement(AGE, account.person.Age),
                                                new XElement(AMOUNT_OF_MONEY, account.AmountOfMoney));

            XElement xAccLoginInfo = new XElement("Account",
                                new XAttribute(ID, account.ID),
                                new XElement(LOGIN, account.Login),
                                new XElement(PASSWORD, account.Password));

            try
            {
                xAccountDocument = XDocument.Load(ACCOUNTS_INFO_DOCUMENT);
                xAccountDocument.Element(ACCOUNTS_INFO_ELEM).Add(xAccPerson);

                xAccountLoginInfoDocument = XDocument.Load(ACCOUNTS_LOGIN_INFO_DOCUMENT);
                xAccountLoginInfoDocument.Element(ACCOUNTS_LOGIN_INFO_ELEM).Add(xAccLoginInfo);


            }
            catch (Exception)
            {
                xAccountDocument.Add(xAccRootElem);
                xAccRootElem.Add(xAccPerson);

                xAccountLoginInfoDocument.Add(xAccLoginInfoRootElem);
                xAccLoginInfoRootElem.Add(xAccLoginInfo);
            }
            xAccountLoginInfoDocument.Save(ACCOUNTS_LOGIN_INFO_DOCUMENT);
            xAccountDocument.Save(ACCOUNTS_INFO_DOCUMENT);
        }


        /// <summary>
        /// Opening or creating xml file with ATM information
        /// </summary>
        /// <param name="ATM"></param>
        public static void OpenOrCreateXmlAtmFile(IATM ATM)
        {
            XDocument xAtmDocument = new XDocument();
            XElement xAtmRootElement = new XElement(ATM_INFO_ELEM);

            XElement xAtm = new XElement("ATM",
                                         new XAttribute(NUMBER_ATM, ATM.NumberATM),
                                         new XElement(ADRESS, ATM.Adress),
                                         new XElement(AMOUNT_OF_MONEY, ATM.AmountOfMoneyATM));
            try
            {
                xAtmDocument = XDocument.Load(ATM_INFO_DOCUMENT);
                xAtmDocument.Element(ATM_INFO_ELEM).Add(xAtm);
            }
            catch (Exception)
            {
                xAtmRootElement.Add(xAtm);
                xAtmDocument.Add(xAtmRootElement);
            }
            xAtmDocument.Save(ATM_INFO_DOCUMENT);
        }
        /// <summary>
        /// Bank withdrawal method. Change sum in xml file
        /// </summary>
        /// <param name="account">Account where the money is withdrawn from</param>
        /// <param name="amountMoney">Available sum on an account</param>
        /// <param name="tempAmountMoney">Desire sum to withdraw</param>
        /// <returns></returns>
        public static double WithdrawMoneyXml(IAccount account, double amountMoney, double tempAmountMoney)
        {
            XDocument xAccountDocument = XDocument.Load(ACCOUNTS_INFO_DOCUMENT);
            XElement xAccountElement = xAccountDocument.Element(ACCOUNTS_INFO_ELEM);

            foreach (var person in xAccountElement.Elements())
            {
                if (int.Parse(person.Attribute(ID).Value) == account.ID)
                {
                    person.Element(AMOUNT_OF_MONEY).Value = (amountMoney -= tempAmountMoney).ToString();
                    xAccountDocument.Save(ACCOUNTS_INFO_DOCUMENT);
                    break;
                }
            }

            return amountMoney;
        }
        /// <summary>
        /// ATM withdrawal method. Change sum in xml file
        /// </summary>
        /// <param name="ATM">Bankomat where the money is withdrawn from</param>
        /// <param name="amountMoney">Available sum on an ATM</param>
        /// <param name="tempAmountMoney">Desire sum to withdraw</param>
        /// <returns></returns>
        public static double WithdrawMoneyAtmXml(IATM ATM, double amountMoney, double tempAmountMoney)
        {
            XDocument xAtmDocument = XDocument.Load(ATM_INFO_DOCUMENT);
            XElement xAtmElement = xAtmDocument.Element(ATM_INFO_ELEM);
            foreach (var atm in xAtmElement.Elements())
            {
                if (int.Parse(atm.Attribute(NUMBER_ATM).Value) == ATM.NumberATM)
                {
                    atm.Element(AMOUNT_OF_MONEY).Value = (amountMoney -= tempAmountMoney).ToString();
                    xAtmDocument.Save(ATM_INFO_DOCUMENT);
                    break;
                }
            }
            return amountMoney;
        }
        /// <summary>
        /// Load Accounts list
        /// </summary>
        /// <returns></returns>
        public static List<IAccount> LoadXmlListAccount()
        {
            List<IAccount> accountsList = new List<IAccount>();
            try
            {
                XDocument xAccountDocument = XDocument.Load(ACCOUNTS_INFO_DOCUMENT);
                XDocument xAccountLoginDocument = XDocument.Load(ACCOUNTS_LOGIN_INFO_DOCUMENT);

                XElement persons = xAccountDocument.Element(ACCOUNTS_INFO_ELEM);
                XElement accounts = xAccountLoginDocument.Element(ACCOUNTS_LOGIN_INFO_ELEM);

                foreach (var person in persons.Elements())
                {
                    Account tempAccount = new Account();
                    tempAccount.ID = int.Parse(person.Attribute(ID).Value);
                    tempAccount.person.Name = person.Element(NAME).Value.ToString();
                    tempAccount.person.SurName = person.Element(SURNAME).Value.ToString();
                    tempAccount.person.Age = int.Parse(person.Element(AGE).Value);
                    tempAccount.AmountOfMoney = int.Parse(person.Element(AMOUNT_OF_MONEY).Value);

                    foreach (var account in accounts.Elements())
                    {
                        if (tempAccount.ID == int.Parse(account.Attribute(ID).Value))
                        {
                            tempAccount.Login = account.Element(LOGIN).Value.ToString();
                            tempAccount.Password = account.Element(PASSWORD).Value.ToString();
                            break;
                        }
                    }
                    accountsList.Add(tempAccount);
                }
                // return list
                return accountsList;
            }
            catch (Exception)
            {
                return accountsList;
            }
        }
        /// <summary>
        /// Load ATM list
        /// </summary>
        /// <returns></returns>
        public static List<IATM> LoadXmlListAtm()
        {
            List<IATM> listATM = new List<IATM>();
            try
            {
                XDocument xAtmDocument = XDocument.Load(ATM_INFO_DOCUMENT);
                XElement ATM = xAtmDocument.Element(ATM_INFO_ELEM);

                foreach (var item in ATM.Elements())
                {
                    ATM atm = new ATM();

                    atm.NumberATM = int.Parse(item.Attribute(NUMBER_ATM).Value);
                    atm.Adress = item.Element(ADRESS).Value.ToString();
                    atm.AmountOfMoneyATM = int.Parse(item.Element(AMOUNT_OF_MONEY).Value);
                    listATM.Add(atm);
                }

                ///retur list ATM
                return listATM;
            }
            catch (Exception)
            {
                return listATM;
            }
        }
    }
}
