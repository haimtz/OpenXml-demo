﻿using System.Collections.Generic;
using System.Linq;

namespace DemoXml
{
    class Program
    {
        static void Main(string[] args)
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
            var document = excel.Document(@"D:\demo.xlsx");
            excel.AddSheet(document.WorkbookPart, "Phone Book");
            var sheet = document.WorkbookPart.WorksheetParts.First().Worksheet;

            excel.AddRow(sheet, true, "Name", "Family", "Phone");

            foreach (var person in persons)
            {
                excel.AddRow(sheet, false ,person.Name, person.Family, person.Phone);
            }
            sheet.Save();
            document.Close();

            System.Diagnostics.Process.Start(@"D:\demo.xlsx");
        }
    }
}
