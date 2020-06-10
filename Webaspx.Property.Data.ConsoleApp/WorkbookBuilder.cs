using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace Webaspx.Property.Data.ConsoleApp
{
    public class WorkbookBuilder
    {
        private IWorkbook _workbook;
        private IDictionary<string, ISheet> _sheets; 
        
        public WorkbookBuilder()
        {
            _workbook = new HSSFWorkbook();
            _sheets = new Dictionary<string, ISheet>();
        }

        public WorkbookBuilder AddSheet(string sheetName) {
            _sheets.Add(sheetName, _workbook.CreateSheet(sheetName));
            return this;
        }

        public WorkbookBuilder AddPropertyHeaders(string sheetName) 
        {
            if(!_sheets.ContainsKey(sheetName)) throw new ArgumentOutOfRangeException(nameof(sheetName), "Sheet does not exist");

            var row = _sheets[sheetName].CreateRow(0);

            row.CreateCell(0).SetCellValue("WaxID");
            row.CreateCell(1).SetCellValue("PostCode");
            row.CreateCell(2).SetCellValue("Streetname");
            row.CreateCell(3).SetCellValue("Address");
            row.CreateCell(4).SetCellValue("Town");

            return this;
        }

        public string CreateWorkbook(string dataFilePath) 
        {

            var resultFilePath = GetResultFilePath(dataFilePath);

            if (File.Exists(resultFilePath)) File.Delete(resultFilePath);

            using(var fileStream = File.Create(resultFilePath)) 
            {
                _workbook.Write(fileStream);
            }

            _workbook = new HSSFWorkbook();
            _sheets = new Dictionary<string, ISheet>();

            return resultFilePath;

        }

        private static string GetResultFilePath(string dataFilePath)
        {
            var resultFileName = "result.xls";

            var path = Regex.Match(dataFilePath, "(.*)(\\/.*\\.xls)");

            return path.Groups[1].Value + "/" + resultFileName;
        }
    }
}