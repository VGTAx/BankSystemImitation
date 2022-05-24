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
        public Person CreatePerson()
        {
            Age = InitializationHelper.IntInit("age");
            Name = InitializationHelper.StringInIt("name");
            SurName = InitializationHelper.StringInIt("surname");
            return new Person(Age, Name, SurName);
        }

        public static bool CheckAge(Person person)
        {
            Type? type = Type.GetType("BankSystem.Person, BankSystem");
            object[] attributes = type.GetCustomAttributes(false);
            foreach (var attr in attributes)
            {
                if(attr is CheckAgeAttribute checkAge)
                {
                    if (checkAge.Age < person.Age)
                        return true;
                    MessageInformant.ErrorOutput("Minors cannot open a bank account!");
                    return false;
                }
            }
            return true;
        }
    }
}
