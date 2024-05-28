using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace XMLassignment
{
    public class FoodDetails
    {
        public string Name { get; set; }
        public double Price { get; set; }
        public string Description { get; set; }
        public int Calories { get; set; }
        public Food Food { get; set; } 
        public FoodDetails(Food food)
        {
            Food = food;
        }

    }
}