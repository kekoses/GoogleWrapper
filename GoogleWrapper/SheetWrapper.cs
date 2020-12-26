using GoogeWrapperLibrary;
using System;

namespace GoogleWrapper
{
    class Program
    {
            static void Main(string[] args)
            {
                    var wrapper = new SheetWrapper("1H6uZOYrVOoq7x2WBrk4IPEVpjZEIkvZ6qK_SKaetNb4");
                    var cell = wrapper.PutCell(5, 2, "Максим Денисов");
                    if(cell.Value.Equals("Максим Денисов"))
                    {
                        Console.WriteLine(cell.Value);
                    }
                    else Console.WriteLine("Не работает");
                    Console.ReadLine();
            }
    }
}
