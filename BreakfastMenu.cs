using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
namespace XMLassignment
{
    public class BreakfastMenu
    {
        private List<Food> _foodsList;

        public List<Food> FoodsList
        {
            get
            {
                if (_foodsList == null)
                {
                    _foodsList = new List<Food>();
                }
                return _foodsList;
            }
        }


        public void PrintMenu()
        {
            foreach (var food in _foodsList)
            {
                Console.WriteLine("Order ID: " + food.OrderID);
                Console.WriteLine("Name: " + food.Detail.Name);
                Console.WriteLine("Price: $" + food.Detail.Price);
                Console.WriteLine("Description: " + food.Detail.Description);
                Console.WriteLine("Calories: " + food.Detail.Calories);
                Console.WriteLine("--------------------------------------------------------------------------------------");
            }
        }


    }

}