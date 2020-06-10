using System;
using System.Collections.Generic;
using System.IO;
using Npoi.Mapper;
using NPOI.SS.UserModel;

namespace Webaspx.Property.Data.ConsoleApp
{
    public static class SpreadSheetHandler
    {
        public static IEnumerable<RowInfo<Model.Property>> ReadRows(string path, int sheetIndex = 0)
        {
            IWorkbook workbook;

            var properties = new List<Model.Property>();

            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                workbook = WorkbookFactory.Create(file);
            }

            var importer = new Mapper(workbook);

            try
            {
                return importer.Take<Model.Property>(sheetIndex);
            }
            catch (ArgumentException)
            {
                throw new ArgumentOutOfRangeException(nameof(sheetIndex), "Sheet does not exist.");
            }
        }

        public static bool WriteRows(string path, string sheetName, IEnumerable<Model.Property> properties)
        {
            var exporter = new Mapper(path);

            try 
            {
                exporter.Put(properties, sheetName, true);
                exporter.Save(path);
                return true;
            }
            catch(IOException ex) 
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }
    }
}