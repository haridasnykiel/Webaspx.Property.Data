using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Npoi.Mapper;
using NPOI.SS.UserModel;
using NUnit.Framework;
using Webaspx.Property.Data.ConsoleApp;
using Webaspx.Property.Data.ConsoleApp.Model;

namespace Tests
{
    public class ReadDataFileTests
    {
        private string _path = "TestData/testdata.xls";

        [Test]
        public void ReadTestFile_SheetCount()
        {
            IWorkbook workbook;

            var properties = new List<Property>();

            using (FileStream file = new FileStream(_path, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(file);
            }

            Assert.AreEqual(workbook.NumberOfSheets, 2);
        }

        [Test]
        public void ReadTestFile_SheetsRowCount()
        {
            IWorkbook workbook;

            var properties = new List<Property>();

            using (FileStream file = new FileStream(_path, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(file);
            }
            var importer = new Mapper(workbook);
            var sheetOneItems = importer.Take<Property>(0);
            var sheetTwoItems = importer.Take<Property>(1);

            var sheetOneRowCount = sheetOneItems.Count();
            var sheetTwoRowCount = sheetTwoItems.Count();
            
            Assert.AreEqual(sheetOneRowCount, 30);
            Assert.AreEqual(sheetTwoRowCount, 20);
        }

        [Test]
        public void SpreadSheetHandler_ReadRows_SheetOne() 
        {
            var rows  = SpreadSheetHandler.ReadRows(_path, 0);
            
            var count = rows.Count();

            Assert.AreEqual(count, 30);
        }

        [Test]
        public void SpreadSheetHandler_ReadRows_SheetTwo()
        {
            var rows = SpreadSheetHandler.ReadRows(_path, 1);

            var count = rows.Count();

            Assert.AreEqual(count, 20);
        }

        [Test]
        public void SpreadSheetHandler_ReadRows_SheetDoesNotExist()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => SpreadSheetHandler.ReadRows(_path, 2));
        }

        [Test]
        public void SpreadSheetHandler_ReadRows_ReturnsProperties() {
            var rows = SpreadSheetHandler.ReadRows(_path, 0);
            var enumerator = rows.GetEnumerator();
            enumerator.MoveNext();

            var property = enumerator.Current.Value;

            Assert.AreEqual(property.WaxId, 430);
            Assert.AreEqual(property.Address, "Salvation Army Halls Arkwright Walk , Fictional Town");
            Assert.AreEqual(property.Postcode, "XX2 2HN");
            Assert.AreEqual(property.StreetName, "Arkwright Walk North");
        }
    }
}