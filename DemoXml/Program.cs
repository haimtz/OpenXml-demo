using System.Collections.Generic;
using System.Linq;

namespace DemoXml
{
    class Program
    {
        static void Main(string[] args)
        {
            Animals();
            PhoneBook();
        }

        public static void Animals()
        {
            var animals = new List<Animal>
            {
                new Animal {Name = "Kitty", Type = "Cat", Age = 2},
                new Animal {Name = "Robert", Type = "Dog", Age = 7},
                new Animal {Name = "Chippopo", Type = "Monkey", Age = 2},
            };

            var excel = new OpenXmlWrapper();
            var document = excel.Document("Animals.xlsx");
            var sheet = excel.AddSheet(document.WorkbookPart, "My Animals");

            excel.AddRow(sheet, true, "Animal name", "Animal Type", "Age");

            foreach (var animal in animals)
            {
                excel.AddRow(sheet, false, animal.Name, animal.Type, animal.Age);
            }

            document.Close();

            System.Diagnostics.Process.Start("Animals.xlsx");
        }

        public static void PhoneBook()
        {
            var persons = new List<Person>
            {
                new Person {Name = "A1", Family = "B", Phone = "C"},
                new Person {Name = "A2", Family = "B", Phone = "C"},
                new Person {Name = "A3", Family = "B", Phone = "C"},
                new Person {Name = "A4", Family = "B", Phone = "C"},
                new Person {Name = "A5", Family = "B", Phone = "C"}
            };

            var excel = new OpenXmlWrapper();
            var document = excel.Document("demo.xlsx");
            excel.AddSheet(document.WorkbookPart, "Phone Book");
            var sheet = document.WorkbookPart.WorksheetParts.First().Worksheet;

            excel.AddRow(sheet, true, "Name", "Family", "Phone");

            foreach (var person in persons)
            {
                excel.AddRow(sheet, false, person.Name, person.Family, person.Phone);
            }
            //sheet.Save();
            document.Close();

            System.Diagnostics.Process.Start("demo.xlsx");
        }
    }
}
