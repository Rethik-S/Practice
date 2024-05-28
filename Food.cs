using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace XMLassignment
{
    public class Food
    {
        public string OrderID { get; set; }
        public BreakfastMenu Menu { get; set; }
        public FoodDetails Detail { get; set; }

        public Food(BreakfastMenu menu)
        {
            Menu = menu;
        }

    }

}