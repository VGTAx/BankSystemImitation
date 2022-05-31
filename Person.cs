using InitHelperInformatMessage;
using Attributes;

namespace BankSystem
{
    [CheckAge(Age = 18)]
    public sealed class Person
    {
        public int Age { get; set; }
        public string Name { get; set; }
        public string SurName { get; set; }
        public void Info()
        {
            Console.WriteLine($"Name - {Name}\nSurname - {SurName}\nAge - {Age}");
        }
        public Person( int age, string name, string surName)
        {
            Age = age;
            Name = name;
            SurName = surName;
        }
        public Person() { }

        public static bool CheckAge(Person person)
        {
            Type? type = Type.GetType("BankSystem.Person, BankSystem");
            object[] attributes = type.GetCustomAttributes(false);
            foreach (var attr in attributes)
            {
                if(attr is CheckAgeAttribute checkAge)
                {
                    if (checkAge.Age <= person.Age)
                        if (100 <= checkAge.Age)
                            return true;
                        else
                        {
                            MessageInformant.ErrorOutput($"Incorrect value! Age must be" +
                                    $" more {checkAge.Age} and less {100}");
                            return false;
                        }  
                    MessageInformant.ErrorOutput("Minors cannot open a bank account!");
                    return false;
                }
            }
            return true;
        }
    }
}
