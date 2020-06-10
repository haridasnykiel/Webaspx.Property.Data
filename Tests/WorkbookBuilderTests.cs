using System;
using System.IO;
using NPOI.HSSF.UserModel;
using NUnit.Framework;
using Webaspx.Property.Data.ConsoleApp;

namespace Tests
{
    public class WorkbookBuilderTests
    {
        private WorkbookBuilder _workbookBuilder;
        private const string _dataFilePath = "/Users/Hari/Desktop/InteviewChallenges/Webaspx/TestData/testdata.xls";

        public WorkbookBuilderTests()
        {
            _workbookBuilder = new WorkbookBuilder();
        }

         [Test]
        public void CreateWorkBook_WithOneSheet() 
        {
            var sheetName = "sheet1";

            var resultPath = _workbookBuilder
                                .AddSheet(sheetName)
                                .CreateWorkbook(_dataFilePath);

            using(var fs = new FileStream(resultPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)) 
            {
                var book = new HSSFWorkbook(fs);

                Assert.AreEqual(book.NumberOfSheets, 1);
            }
        }

        [Test]
        public void CreateWorkBook_WithTwoSheet()
        {
            var sheetOne = "sheet1";
            var sheetTwo = "sheet2";

            var resultPath = _workbookBuilder
                                .AddSheet(sheetOne)
                                .AddSheet(sheetTwo)
                                .CreateWorkbook(_dataFilePath);

            using (var fs = new FileStream(resultPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var book = new HSSFWorkbook(fs);

                Assert.AreEqual(book.NumberOfSheets, 2);
            }
        }

        [Test]
        public void CreateWorkBook_WithPropertyHeaders()
        {
            var sheetOne = "sheet1";

            var resultPath = _workbookBuilder
                                .AddSheet(sheetOne)
                                .AddPropertyHeaders(sheetOne)
                                .CreateWorkbook(_dataFilePath);

            using (var fs = new FileStream(resultPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                var book = new HSSFWorkbook(fs);

                var sheet = book.GetSheet(sheetOne);

                var header = sheet.GetRow(0);

                var cellsInHeaderRow = header.Cells;

                Assert.AreEqual(cellsInHeaderRow[0].StringCellValue, "WaxID");
                Assert.AreEqual(cellsInHeaderRow[1].StringCellValue, "PostCode");
                Assert.AreEqual(cellsInHeaderRow[2].StringCellValue, "Streetname");
                Assert.AreEqual(cellsInHeaderRow[3].StringCellValue, "Address");
                Assert.AreEqual(cellsInHeaderRow[4].StringCellValue, "Town");
            }
        }

        [Test]
        public void AddPropertyHeaders_ThrowsWhenSheetDoesNotExist() 
        {
            var sheetOne = "sheet1";

            Assert.Throws<ArgumentOutOfRangeException>(() => _workbookBuilder.AddPropertyHeaders(sheetOne));
        }
    }
}