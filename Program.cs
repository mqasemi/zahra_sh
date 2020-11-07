using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;

namespace zahra_shastcom
{
    class Program
    {
        static void Main(string[] args)
        {
            string baseFileDirectory = Directory.GetCurrentDirectory() + @"\DB\Base\";
            string workDirectory = Directory.GetCurrentDirectory() + @"\DB\work\";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var baseFilePath = FileInputUtil.GetFileInfo(baseFileDirectory, Directory.GetFiles(baseFileDirectory).FirstOrDefault() ).FullName;
            FileInfo existingBaseFile = new FileInfo(baseFilePath);
            ExcelPackage basePackage = new ExcelPackage(existingBaseFile);
            var workFiles = Directory.GetFiles(workDirectory).ToList();
            ExcelWorksheet baseWorksheet = basePackage.Workbook.Worksheets[0];
            var basecolum = baseWorksheet.GetHeaderColumns();
            var basePcodeIndex = Array.FindIndex(basecolum, c => c.Equals("کدپرسنلی"));
            var baseLnameIndex = Array.FindIndex(basecolum, c => c.Equals("نام خانوادگی"));
            var baseFnameIndex = Array.FindIndex(basecolum, c => c.Equals("نام"));
            var basefnameLnameIndex = Array.FindIndex(basecolum, c => c.Equals("نام و نام خانوادگی"));


            foreach (var wfile in workFiles)
            {
                //var workFilePath = FileInputUtil.GetFileInfo(workDirectory, wfile).FullName;
                FileInfo existingWorkFile = new FileInfo(wfile);
                ExcelPackage workPackage = new ExcelPackage(existingWorkFile);
                ExcelWorksheet workWorkSheet = workPackage.Workbook.Worksheets[0];

                var workColumn = workWorkSheet.GetHeaderColumns();
                var workPcodeIndex = Array.FindIndex(workColumn, c => c.Equals("کدپرسنلی"));
                var workLnameIndex = Array.FindIndex(workColumn, c => c.Equals("نام خانوادگی"));
                var workFnameIndex = Array.FindIndex(workColumn, c => c.Equals("نام"));
                var workfnameLnameIndex = Array.FindIndex(workColumn, c => c.Equals("نام و نام خانوادگی"));

                if (workPcodeIndex > 0 && (workfnameLnameIndex > 0 || (workLnameIndex > 0 && workFnameIndex > 0)))
                {
                    for (int row = 2; row <= workWorkSheet.Dimension.End.Row; row++)
                    {
                        Int64 pcode = workWorkSheet.Cells[row, workPcodeIndex + 1].Value != null ? Int64.Parse(workWorkSheet.Cells[row, workPcodeIndex + 1].Value.ToString()) : -1;

                        if (pcode > 0)
                        {
                            var query1 = (from cell in baseWorksheet.Cells[2, basePcodeIndex + 1, baseWorksheet.Dimension.End.Row, basePcodeIndex + 1]
                                          where
                                              cell.Value.ToString() == pcode.ToString()
                                          select cell).FirstOrDefault();

                            string baseFnameLname = baseWorksheet.Cells[query1.Start.Row, basefnameLnameIndex + 1].Value.ToString().Trim();
                            if (baseFnameLname != null)
                            {
                                var splitNames = baseFnameLname.Split(' ');

                                baseFnameLname = string.Join(" ", splitNames.Where(c => c != "").ToArray());
                            }
                            string workFnameLname = "";
                            if (workfnameLnameIndex > 0)
                            {
                                workFnameLname = workWorkSheet.Cells[row, workfnameLnameIndex + 1].Value.ToString();
                                if (workFnameLname != null)
                                {
                                    var splitNames = workFnameLname.Split(' ');
                                    workFnameLname = string.Join(" ", splitNames.Where(c => c != "").ToArray());
                                }

                            }
                            else
                            {
                                var wfnamelname = workWorkSheet.Cells[row, workFnameIndex + 1].Value.ToString()+" "+workWorkSheet.Cells[row, workLnameIndex + 1].Value.ToString();
                                var wlfnames = wfnamelname.Split(' ');
                                workFnameLname = string.Join(" ", wlfnames.Where(c => c != "").ToArray());

                            }

                            if (!baseFnameLname.Equals(workFnameLname))
                            {
                                workWorkSheet.Cells[row, workWorkSheet.Dimension.End.Column + 1].Value = baseFnameLname;
                                pcode = -1;
                                //
                            }
                        }


                        if (pcode < 0)
                        {
                            // workWorkSheet.Cells[row, workWorkSheet.Dimension.Start.Column, row, workWorkSheet.Dimension.End.Column].Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml("#ff0000bf"));
                            workWorkSheet.Row(row).Style.Fill.PatternType = ExcelFillStyle.Solid;
                            workWorkSheet.Row(row).Style.Fill.BackgroundColor.SetColor(Color.Pink);
                        }


                       

                    }

                }

                FileOutputUtil.OutputDir = new DirectoryInfo(workDirectory);
                var xlFile = FileOutputUtil.GetFileInfo("result"+existingWorkFile.Name);
                // sheet.Column(idx).Style.Numberformat.Format = "mm-dd-yy";
                workPackage.SaveAs(xlFile); 

            }


			Console.WriteLine();
			Console.WriteLine("Read workbook sample complete");
            Console.ReadLine();
		}


       

       
        
    }
}
