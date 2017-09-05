using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace Get_data_and_manipulate_it_and_export_it
{
    class Program
    {
        static void Main(string[] args)
        {
            int i, j;
            //double fValue;
            Random rand = new Random();
            double[] doubledata = new double[] { rand.Next(10),rand.NextDouble(),2.3,4.5,6.9};

            using (SLDocument sl = new SLDocument())
            {
                sl.RenameWorksheet(SLDocument.DefaultFirstSheetName, "Random");

                for (i = 1; i <= 16; ++i)
                {
                    for (j = 1; j <= 6; ++j)
                    {
                        switch (rand.Next(5))
                        {
                            case 0:
                            case 1:
                                sl.SetCellValue(i, j, doubledata[rand.Next(doubledata.Length)]);
                                break;
                            case 2:
                            case 3:
                                sl.SetCellValue(i, j, rand.NextDouble() * 1000.0 + 350.0);
                                break;

                            case 4:
                                if (rand.NextDouble() < 0.5)
                                {
                                    sl.SetCellValueNumeric(i, j, "3.1415926535898");
                                }
                                else
                                {
                                    sl.SetCellValueNumeric(i, j, "2.7182818284590");
                                }
                                break;
                        }
                    }
                }
                //sl.Filter("A1", "F1");
                //sl.FlattenAllSharedCellFormula();
                //sl.DrawBorderGrid 
                //sl.SetCellValue("C6", "This is at C6!");
                //sl.SetCellValue(SLConvert.ToCellReference(7, 4), string.Format("=SUM({0})", SLConvert.ToCellRange(2, 1, 2, 20)));
                //sl.SetCellValue("C8", new DateTime(3141, 5, 9));
                //SLStyle style = sl.CreateStyle();
                //style.FormatCode = "d-mmm-yyyy";
                //sl.SetCellStyle("C8", style);
                //for (int i = 1; i <= 20; ++i) sl.SetCellValue(2, i, i);
          
                sl.SaveAs("Miscellaneous1.xlsx");
                SLDocument tl = new SLDocument("Miscellaneous1.xlsx", "Sheet");
                //以下两个方法功能一致
                tl.SetCellValue("G1", "=SUM(A1:F2)");
                tl.SetCellValue(SLConvert.ToCellReference(2, 7), string.Format("=SUM({0})", SLConvert.ToCellRange(1, 1, 2, 6)));
                tl.SetCellValue("G3", "=A2-A3");
                //tl.SetCellValueNumeric("G6", "=AVERAGE(A1:F2)");
                //tl.SetCellValue("G2", string.Format(" =SUM({0})", SLConvert.ToCellRange(1, 1, 2, 6)));
                //以下两个方法功能一致
                tl.SetCellValue("G4", StringValue.ToString("=AVERAGE(A1:F2)"));
                tl.SetCellValue("G5", String.Format("=AVERAGE({0})",SLConvert.ToCellRange(1,1,2,6)));
                tl.SetCellValue("G6", "So this is the random number table");
                tl.SetCellValue("G7", "The time is ");
                tl.SetCellValue("H7", DateTime.Now.ToString());//获取当前日期和时间
                tl.AddWorksheet("Secret");//添加工作表
                tl.SelectWorksheet("Random");//选择Random工作表
                tl.RenameWorksheet("Secret", "again");//重命名Secret
                SLStyle style = tl.CreateStyle();//设置单元格格式
                style.SetFont("Impact", 24);
                style.Font.Underline = UnderlineValues.Single;
                tl.SetCellStyle(1, 7, style);
                tl.SetCellStyle("G6", style);
                //设置EXCEL属性
                tl.DocumentProperties.Creator = "ZhouL";
                tl.DocumentProperties.ContentStatus = "Secret";
                tl.DocumentProperties.Title = "Random number table";
                tl.DocumentProperties.Description = "Get data and manipulate it and export it";
                tl.SaveAs("MiscellaneousModified.xlsx");

            }
            Console.WriteLine("End of program");
            Console.ReadLine();
        }
    }
}
