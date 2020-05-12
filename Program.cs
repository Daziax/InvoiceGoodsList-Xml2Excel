using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text;


namespace FrameWork发票
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelOp eo = new ExcelOp();
            //Console.WriteLine("请输入税控盘持有人姓名:");
            
           
            Console.WriteLine("请讲文件拖入此处(文件的名称格式为：xx商品编码，xx为税控盘名称！且软件运行后生成的excel文件在 《文档》 文件夹中)\n");
            string fullname = Console.ReadLine();

            Console.WriteLine("\n正在转换中，请耐心等待。。。");
            string name = fullname.Substring(fullname.IndexOf(".") - 6,6);
            eo.CreateExcelFile(name);
            object Nothing = System.Reflection.Missing.Value;
            Excel.Application app = new Excel.Application();
            app.Visible = false;

            //string item = eo.Read(string.Format("E://个人文件//妈妈//{0}商品编码.txt",name));
            string item = eo.Read(fullname);
            int i = item.IndexOf("<BMXX>");
            int i_ = item.IndexOf("</BMXX>");
            try
            {
                while (i != -1)
                {
                    Excel.Workbook mybook = app.Workbooks.Open(name, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing, Nothing);
                    Excel.Worksheet mysheet = (Excel.Worksheet)mybook.Worksheets[1];
                    int i1 = item.IndexOf("<SPMC>");
                    int i1_ = item.IndexOf("</SPMC>");
                    int i2 = item.IndexOf("<SPBM>");
                    int i2_ = item.IndexOf("</SPBM>");
                    int i3 = item.IndexOf("<JM>");
                    int i3_ = item.IndexOf("</JM>");
                    int i4 = item.IndexOf("<SPBMJC>");
                    int i4_ = item.IndexOf("</SPBMJC>");
                    int i5 = item.IndexOf("<ZZSSL>");
                    int i5_ = item.IndexOf("</ZZSSL>");
                    int i6 = item.IndexOf("<GGXH>");
                    int i6_ = item.IndexOf("</GGXH>");
                    int i7 = item.IndexOf("<JLDW>");
                    int i7_ = item.IndexOf("</JLDW>");
                    int i8 = item.IndexOf("<KYSL>");
                    int i8_ = item.IndexOf("</KYSL>");
                    int i9 = item.IndexOf("<HSBZ>");
                    int i9_ = item.IndexOf("</HSBZ>");
                    int i10 = item.IndexOf("<YH>");//优惠政策类型
                    int i10_ = item.IndexOf("</YH>");
                    int i11 = item.IndexOf("<SYPC>");//免税类型
                    int i11_ = item.IndexOf("</SYPC>");

                    string c1 = item.Substring(i1 + 6, i1_ - i1 - 6);
                    string c2 = item.Substring(i2 + 6, i2_ - i2 - 6);
                    string c3 = item.Substring(i3 + 4, i3_ - i3 - 4);
                    string c4 = item.Substring(i4 + 8, i4_ - i4 - 8);
                    string c5 = item.Substring(i5 + 7, i5_ - i5 - 7);
                    string c6 = item.Substring(i6 + 6, i6_ - i6 - 6);
                    string c7 = item.Substring(i7 + 6, i7_ - i7 - 6);
                    string c8 = item.Substring(i8 + 6, i8_ - i8 - 6);
                    string c9 = item.Substring(i9 + 6, i9_ - i9 - 6);
                    string c10 = item.Substring(i10 + 4, i10_ - i10 - 4);
                    string c11 = item.Substring(i11 + 6, i11_ - i11 - 6);
                    eo.WriteToExcel("全部商品", c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11, mybook, mysheet, Nothing, app);
                    item = item.Remove(0, i_ + 7);
                    i = item.IndexOf("<BMXX>");
                    i_ = item.IndexOf("</BMXX>");
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
            finally
            {
                Console.WriteLine("\n转换完成！请在‘我的文档’中查找点击！输入回车自动退出。");
                Console.ReadKey();
                app.Quit();
            }



        }

    }
    class ExcelOp
    {
        public string Read(string path)
        {
            StreamReader sr = new StreamReader(path, Encoding.UTF8);
            return sr.ReadToEnd();
        }

        internal void CreateExcelFile(string FileName)
        {
            //create
            object Nothing = System.Reflection.Missing.Value;
            var app = new Excel.Application();
            app.Visible = false;
            Excel.Workbook workBook = app.Workbooks.Add(Nothing);
            Excel.Worksheet worksheet = (Excel.Worksheet)workBook.Sheets[1];
            
            worksheet.Name = FileName;
            //headline
            worksheet.Cells[1, 1] = "名称";
            worksheet.Cells[1, 2] = "编码";
            //worksheet.Cells[1, 2].numberFormatting = "@";
            worksheet.Cells[1, 3] = "简码";
            worksheet.Cells[1, 4] = "税收分类简称";
            worksheet.Cells[1, 5] = "税率";
            //worksheet.Cells[1, 5].numberFormatting = "formatnumber";
            worksheet.Cells[1, 6] = "规格/厂牌";
            worksheet.Cells[1, 7] = "计量单位";
            worksheet.Cells[1, 8] = "适用税率";
            worksheet.Cells[1, 9] = "含税标志";
            worksheet.Cells[1, 10] = "优惠政策类型";
            worksheet.Cells[1, 11] = "免税类型";

            worksheet.Columns[2].NumberFormatLocal = "@";//设置第二列为 文本格式

            worksheet.SaveAs(FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing);
            workBook.Close(false, Type.Missing, Type.Missing);


            app.Quit();

        }
        internal void WriteToExcel(string excelName, string c1, string c2, string c3, string c4, string c5, string c6, string c7, string c8, string c9, string c10, string c11, Excel.Workbook mybook, Excel.Worksheet mysheet, object Nothing, Excel.Application app)
        {
            //open
            mysheet.Activate();
            //get activate sheet max row count
            int maxrow = mysheet.UsedRange.Rows.Count + 1;
            mysheet.Cells[maxrow, 1] = c1;
            mysheet.Cells[maxrow, 2] = c2;
            mysheet.Cells[maxrow, 3] = c3;
            mysheet.Cells[maxrow, 4] = c4;
            mysheet.Cells[maxrow, 5] = c5;
            mysheet.Cells[maxrow, 6] = c6;
            mysheet.Cells[maxrow, 7] = c7;
            mysheet.Cells[maxrow, 8] = c8;
            mysheet.Cells[maxrow, 9] = c9;
            mysheet.Cells[maxrow, 10] = c10;
            mysheet.Cells[maxrow, 11] = c11;
            mybook.Save();
            mybook.Close(false, Type.Missing, Type.Missing);
            mybook = null;
        }
    }
}

