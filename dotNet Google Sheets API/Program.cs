using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace dotNet_Google_Sheets_API
{
    class Program
    {
        static void Main(string[] args)
        {
            string sheetname = "";
            GoogleSheetsAPI GSAPI = new GoogleSheetsAPI();
            IList<IList<object>> Values = GSAPI.GetValuesInSheet(sheetname);
            try
            {
                foreach (IList<object> row in Values)
                {
                    string A5 = Convert.ToString(row[0]);
                    string A6 = Convert.ToString(row[1]);
                    Console.WriteLine("{0} , {1}", A5, A6);

                }
            }
            catch
            {
                Console.WriteLine("Somthing wrong.");
            }
            Console.ReadLine();
        }
    }
}
