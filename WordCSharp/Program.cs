using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace WordCSharp
{
    class Program
    {
        static void Main(string[] args)
        {
            Application app = new Application();
            Document test = app.Documents.Add(Visible: true);
            Range r = test.Range();
            r.Text = "Hello, Word";
            //r.Bold = 20;

            Table tab = test.Tables.Add(r, 5, 5);
            tab.Borders.Enable = 1;

            foreach (Row row in tab.Rows)
            {
                foreach (Cell cell in row.Cells)
                {
                    if (cell.RowIndex == 1)
                    {
                        cell.Range.Text = "Колонка " + cell.ColumnIndex.ToString();
                        cell.Range.Bold = 1;
                        cell.Range.Font.Name = "Arial";
                        cell.Range.Font.Size = 14;

                        cell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                    }
                    else
                    {
                        //cell.Range.Text = (cell.RowIndex - 2 + cell.ColumnIndex).ToString();
                        cell.Range.Text = "Hello, Word";
                    }
                }
            }

            test.Save();
            app.Documents.Open("C:\\Users\\Илья\\Desktop\\Doc2.docx");
            Console.ReadKey();
            try
            {
                test.Close();
                app.Quit();
            }
            catch (Exception a)
            {
                Console.WriteLine(a.Message);
            }
            Console.ReadKey();

        }
    }
}
