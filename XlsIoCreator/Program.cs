using FizzWare.NBuilder;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XlsIoCreator
{
    class Program
    {
        static void Main(string[] args)
        {
            var customers = Builder<Customer>.CreateListOfSize(100)
            .All()
                .With(c => c.FirstName = Faker.Name.First())
                .With(c => c.LastName = Faker.Name.Last())
                .With(c => c.EmailAddress = Faker.Internet.Email())
                .With(c => c.TelephoneNumber = Faker.Phone.Number())
            .Build().ToList();

            Random rnd = new Random();
            for (int i = 0; i < customers.Count; i++)
            {
                var check = rnd.Next(1, 3);
                if (check == 2)
                {
                    int year = rnd.Next(1950, 1999);
                    int month = rnd.Next(1, 12);
                    int day = rnd.Next(1, 28);
                    customers[i].DateOfBirth = new DateTime(year, month, day);
                    customers[i].Salary = rnd.NextDecimal();
                }
                else
                {
                    customers[i].DateOfBirth = null;
                    customers[i].Salary = null;
                }

            }

            var fileName = @"C:\Users\Andrew\Desktop\Test.xlsx";

            if (System.IO.File.Exists(fileName))
                System.IO.File.Delete(fileName);

            var buffer = customers.ToXlsIoBuffer();
            System.IO.File.WriteAllBytes(fileName, buffer);

            System.Diagnostics.Process.Start(fileName);
            //Console.ReadKey();
        }




    }

    public class Customer
    {
        [Display(DisplayName ="Date of Birth")]
        
        public DateTime? DateOfBirth { get; set; }
        public string EmailAddress { get; set; }
        public string FirstName { get; set; }
        public int Id { get; set; }
        public string LastName { get; set; }
        public string TelephoneNumber { get; set; }
        [Currency]
        public decimal? Salary { get; set; }
    }
}
