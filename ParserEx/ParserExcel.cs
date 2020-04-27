using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ParserEx.Models
{
    public static class ParserEx
    {
        public static string PathFile { get; private set; }

        public static void SetPath(string path)
        { 
            PathFile = path;
        }

        public static void BaseParse()
        {
            //считываем данные из Excel файла в двумерный массив
            Excel.Application xlApp = new Excel.Application(); //Excel
            Excel.Workbook xlWB; //рабочая книга              
            Excel.Worksheet xlSht; //лист Excel   
            xlWB = xlApp.Workbooks.Open(); //название файла Excel                                             
            xlSht = xlWB.Worksheets["Лист1"]; //название листа или 1-й лист в книге xlSht = xlWB.Worksheets[1];
            int iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А            
            var arrData = (object[,])xlSht.Range["A1:J" + iLastRow].Value; //берём данные с листа Excel
            //xlApp.Visible = true; //отображаем Excel     
            xlWB.Close(false); //закрываем книгу, изменения не сохраняем
            xlApp.Quit(); //закрываем Excel

            //настройка DataGridView
          
        }
    }
}
