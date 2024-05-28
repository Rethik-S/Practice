using System;
using System.ComponentModel.Design;
using System.Xml;
namespace XMLassignment;

public class Program
{
    public static void Main(string[] args)
    {
        BreakfastMenu menu = ReadFromXml();
        WriteToXml(menu);
    }

    static BreakfastMenu ReadFromXml()
    {
        BreakfastMenu menu = null;

        using (XmlReader reader = XmlReader.Create("FoodCatelog1.xml"))
        {
            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.Element && reader.Name == "breakfast_menu")
                {
                    menu = new BreakfastMenu();
                    MenuReadFromXml(reader, menu);
                    break;
                }
                else if (reader.NodeType == XmlNodeType.Element)
                {
                    Console.WriteLine("Unexpected root element: " + reader.Name);
                    break;
                }

            }

        }

        return menu;

    }
    static void MenuReadFromXml(XmlReader reader, BreakfastMenu menu)
    {
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.Name == "food")
            {
                Food food = new Food(menu);

                food = FoodReadFromXml(reader, food);

                menu.FoodsList.Add(food);
            }

        }
    }
     static Food FoodReadFromXml(XmlReader reader, Food food)
    {


        if (reader.HasAttributes)
        {
            food.OrderID = reader.GetAttribute("orderid");
        }

        food.Detail = FoodDetailsReadFromXml(reader, food);

        return food;


    }

     static FoodDetails FoodDetailsReadFromXml(XmlReader reader, Food food)
    {
        FoodDetails foodDetails = new FoodDetails(food);

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element)
            {
                switch (reader.Name)
                {
                    case "name":
                        reader.Read();
                        foodDetails.Name = reader.Value;
                        break;
                    case "price":
                        reader.Read();
                        double price;
                        if (double.TryParse(reader.Value.TrimStart('$'), out price))
                        {
                            foodDetails.Price = price;
                        }
                        break;
                    case "description":
                        reader.Read();
                        foodDetails.Description = reader.Value;
                        break;
                    case "calories":
                        reader.Read();
                        int calories;
                        if (int.TryParse(reader.Value, out calories))
                        {
                            foodDetails.Calories = calories;
                        }
                        break;
                }
            }
            else if (reader.NodeType == XmlNodeType.EndElement && reader.Name == "food")
            {
                break;
            }

        }

        return foodDetails;
    }


    static void WriteToXml(BreakfastMenu menu)
    {

        using (XmlWriter writer = XmlWriter.Create("output.xml"))
        {
            if (menu != null)
            {
                writer.WriteStartDocument();

                MenuWriteToXml(writer, menu);

                writer.WriteEndDocument();

                Console.WriteLine($"created output.xml");
            }
        }
    }

    static void MenuWriteToXml(XmlWriter writer, BreakfastMenu menu)
    {
        if (menu != null)
        {
            writer.WriteStartElement("breakfast_menu");

            FoodWriteToXml(writer, menu);

            writer.WriteEndElement();
        }

    }


    static void FoodWriteToXml(XmlWriter writer, BreakfastMenu menu)
    {
        foreach (var food in menu.FoodsList)
        {
            writer.WriteStartElement("food");
            if (!string.IsNullOrEmpty(food.OrderID))
            {
                writer.WriteAttributeString("orderid", food.OrderID);
            }
            FoodDetailsWriteToXml(writer, food);
            writer.WriteEndElement();
        }

    }

    static void FoodDetailsWriteToXml(XmlWriter writer, Food food)
    {
        if (food.Detail.Name != null)
        {
            writer.WriteStartElement("name");
            writer.WriteString(food.Detail.Name);
            writer.WriteEndElement();
        }

        if (food.Detail.Price != 0)
        {
            writer.WriteStartElement("price");
            writer.WriteString(food.Detail.Price.ToString("C"));
            writer.WriteEndElement();
        }

        if (food.Detail.Description != null)
        {
            writer.WriteStartElement("description");
            writer.WriteString(food.Detail.Description);
            writer.WriteEndElement();
        }

        if (food.Detail.Calories != 0)
        {
            writer.WriteStartElement("calories");
            writer.WriteString(food.Detail.Calories.ToString());
            writer.WriteEndElement();
        }
    }

}