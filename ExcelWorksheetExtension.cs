using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace zahra_shastcom
{
  public static  class ExcelWorksheetExtension
    {
      public static string[] GetHeaderColumns(this ExcelWorksheet sheet)
      {
          List<string> columnNames = new List<string>();
          foreach (var firstRowCell in sheet.Cells[sheet.Dimension.Start.Row, sheet.Dimension.Start.Column, 1, sheet.Dimension.End.Column])
              columnNames.Add(firstRowCell.Text);
          return columnNames.ToArray();
      }
    }
}
