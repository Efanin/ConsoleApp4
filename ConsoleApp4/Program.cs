using ConsoleApp4;
using System.IO;

Exel exel = new Exel("Goga","C:\\Users\\efani\\source\\repos\\ConsoleApp4\\ConsoleApp4\\base.xlsx");

exel.Add(Console.ReadLine(), Console.ReadLine());
exel.Add(Console.ReadLine(), Console.ReadLine());
exel.Add(Console.ReadLine(), Console.ReadLine());
exel.Close();

