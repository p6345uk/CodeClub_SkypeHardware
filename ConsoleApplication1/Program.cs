using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using lyncx;

namespace ConsoleApplication1
{
    public class Program
    {

        static void Main(string[] args)
        {
            LyncxClient lyncxClient = new LyncxClient();
            lyncxClient.AvailabilityChanged += lyncxClient_AvailabilityChanged;
            lyncxClient.Setup();
            Console.ReadLine();
        }
        public static void lyncxClient_AvailabilityChanged(object sender, AvailabilityChangedEventArgs e)
        {
            try
            {
                Console.WriteLine(e.AvailabilityName);

            }
            catch (Exception ex)
            {
            }
        }
    }

}
