using System;
using System.Collections.Generic;
using System.IO;

namespace Webaspx.Property.Data.ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter file path: ");
            var filePath = Console.ReadLine(); //"/Users/Hari/Desktop/InteviewChallenges/Webaspx/PropertyData/data.xls";
            
            if (!File.Exists(filePath))
            {
                Console.WriteLine("File path does not exist. Please re-enter: ");
                filePath = Console.ReadLine();
            }

            var sheetOne = "sheet1";
            var sheetTwo = "sheet2";

            var sheetOneData = SpreadSheetHandler.ReadRows(filePath, 0);
            var sheetTwoData = SpreadSheetHandler.ReadRows(filePath, 1);

            var sheetOneProperties = PropertyHandler.UpdateTownNames(sheetOneData);
            var sheetTwoProperties = PropertyHandler.UpdateTownNames(sheetTwoData);
            var sheetTwoPropertiesUpdated = PropertyHandler.AddMissingWaxId(sheetTwoProperties, sheetOneProperties);

            var workBookBuilder = new WorkbookBuilder();

            var resultPath = workBookBuilder
                .AddSheet(sheetOne)
                .AddPropertyHeaders(sheetOne)
                .AddSheet(sheetTwo)
                .AddPropertyHeaders(sheetTwo)
                .CreateWorkbook(filePath);

            var updateSheetOne = SpreadSheetHandler.WriteRows(resultPath, sheetOne, sheetOneProperties);
            var updateSheetTwo = SpreadSheetHandler.WriteRows(resultPath, sheetTwo, sheetTwoPropertiesUpdated);

            if (updateSheetOne && updateSheetTwo)
            {
                Console.WriteLine("Town names have been moved to another column \n " +
                                    "Sheet2 WaxIDs have been updated \n " +
                                    $"File location {resultPath}");
            }
            else
            {
                Console.WriteLine("Town names were not moved and result file not produced");
            }

            Console.ReadLine();
        }
    }
}
