using System.Collections.Generic;
using System.IO;
using System.Linq;
using Npoi.Mapper;
using NUnit.Framework;
using Webaspx.Property.Data.ConsoleApp;
using Webaspx.Property.Data.ConsoleApp.Model;

namespace Tests
{
    public class WriteDataToFileTests
    {
        private const string _dataFilePath = "/Users/Hari/Desktop/InteviewChallenges/Webaspx/TestData/testdata.xls";
        private static WorkbookBuilder _workbookBuilder = new WorkbookBuilder();

        [Test]
        public void NpoiMapper_WriteToFile_AddPropertyToFile()
        {
            const string existingFile = "TestData/testdata1.xls";
            const string sheetName = "newSheet";
            if (File.Exists(existingFile)) File.Delete(existingFile);
            File.Copy(_dataFilePath, existingFile);
            var exporter = new Mapper();   
            var property1 = new Property
            {
                WaxId = 1,
                Postcode = "BD13 2AD",
                StreetName = "Bradford Road",
                Address = "23 Bradford Road",
                Town = "Gomersal"
            };
            var property2 = new Property
            {
                WaxId = 2,
                Postcode = "BD13 5AD",
                StreetName = "Bradford Road",
                Address = "25 Bradford Road",
                Town = "Gomersal"
            };
            var properties = new List<Property>();
            properties.Add(property1);
            properties.Add(property2);

            exporter.Put(properties, sheetName);
            exporter.Save(existingFile);

            Assert.IsNotNull(exporter.Workbook);
            Assert.AreEqual(3, exporter.Workbook.GetSheet(sheetName).PhysicalNumberOfRows);
            Assert.AreEqual(property1.Postcode, exporter.Take<Property>(sheetName).First().Value.Postcode);
            Assert.AreEqual(property1.Address, exporter.Take<Property>(sheetName).First().Value.Address);
        }

        [Test]
        public void SpreadSheetHandler_WriteRowsToFile_AddProperties() 
        {
            const string sheetName = "sheet1";

            var expected = new Property
            {
                WaxId = 1,
                Postcode = "BD13 2AD",
                StreetName = "Bradford Road",
                Address = "23 Bradford Road",
                Town = "Gomersal"
            };
            var properties = new List<Property>();
            properties.Add(expected);

            var resultPath = _workbookBuilder
                .AddSheet(sheetName)
                .AddPropertyHeaders(sheetName)
                .CreateWorkbook(_dataFilePath);

            var writeToFile = SpreadSheetHandler.WriteRows(resultPath, sheetName, properties);

            Assert.True(writeToFile);

            var data = SpreadSheetHandler.ReadRows(resultPath);
            var enumerator = data.GetEnumerator();
            enumerator.MoveNext();
            var actual = enumerator.Current.Value;

            Assert.AreEqual(actual.Postcode, expected.Postcode);
            Assert.AreEqual(actual.Address, expected.Address);
            Assert.AreEqual(actual.WaxId, expected.WaxId);
            Assert.AreEqual(actual.StreetName, expected.StreetName);
            Assert.AreEqual(actual.Town, expected.Town);

        }
    }
}