using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using ClosedXML.Excel;
using Dualp.Packages.ExportTo.Models;
using Microsoft.AspNetCore.Mvc;

namespace Dualp.Packages.ExportTo.Extentions
{
   public static  class ExportToExtention
    {
        public static IActionResult ToExcel(this List<KeyValuesProperties> keyValues)
        {
            using var workBook = new XLWorkbook();

            var workSheet = workBook.Worksheets.Add("ExportData");
            var currentRow = 1;
            var currentCol = 1;
            //Add Columns 
            foreach (var keyValue in keyValues)
            {
                workSheet.Cell(currentRow, currentCol).Value = keyValue.DisplayProperty;

                foreach (var value in keyValue.Values)
                {
                    currentRow++;
                    workSheet.Cell(currentRow, currentCol).Value = value;
                }

                currentRow = 1;
                currentCol++;
            }


            using var ms = new MemoryStream();
            workBook.SaveAs(ms);

            

            var array = ms.ToArray();


            ms.Position = 0;

            var nameFile = $"Export-{DateTime.Now:yyyyMMMMdd-hhmm}.xlsx";


            return new FileContentResult(
                array,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                FileDownloadName = nameFile
            };

        }

        public static IActionResult ToCsv(this List<KeyValuesProperties> keyValues)
        {
            var result = new StringBuilder();

            result.Append(string.Join(',', keyValues.Select(x => x.DisplayProperty)));

            var count = keyValues.First().Values.Count;

            for (int i = 0; i < count; i++)
            {
                var line = string.Empty;
                foreach (var keyValue in keyValues)
                {
                    var value = keyValue.Values[i];
                    line = value + ",";
                }

                line.TrimEnd(',');
                result.Append(line);
            }

            var ms = new MemoryStream(Encoding.UTF8.GetBytes(result.ToString()));

            var array = ms.ToArray();


            ms.Position = 0;

            var nameFile = $"Export-{DateTime.Now:yyyyMMMMdd-hhmm}.csv";


            return new FileContentResult(
                array,
                "application/octet-stream")
            {
                FileDownloadName = nameFile
            };
        }
    }
}
