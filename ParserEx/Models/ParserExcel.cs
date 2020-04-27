﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace ParserEx.Models
{
    public static class ParserExcel
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
            xlWB = xlApp.Workbooks.Open(PathFile); //название файла Excel                                             
            xlSht = xlWB.Worksheets["TDSheet"]; //название листа или 1-й лист в книге xlSht = xlWB.Worksheets[1];
            int iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А            
          
            List<int> indexRow = new List<int>();
            for (int i = 1; i < iLastRow; i++)
            {
                var cell = xlSht.Cells[i, 1] as Excel.Range;
                if ( cell.Value == "РАСЧЕТНЫЙ ЛИСТОК ЗА АПРЕЛЬ 2020 (ЗА ПЕРВУЮ ПОЛОВИНУ МЕСЯЦА)")
                {
                    indexRow.Add(i);
                }
            }


            for (int i = 0; i < indexRow.Count; i++)
            {
                dynamic arrData = null;
                string name = "default";
                int maxRow = 0;
                if (i == indexRow.Count - 1)
                {
                    arrData = xlSht.Range["A" + indexRow[i].ToString(), "AI" + iLastRow].Value;
                    var cell = xlSht.Cells[indexRow[i] + 1, 1] as Excel.Range;
                    name = cell.Value;
                    maxRow = iLastRow - indexRow[i]+1;
                }
                else
                {
                    arrData = xlSht.Range["A" + indexRow[i].ToString() , "AI" + indexRow[i + 1]].Value;
                    var cell = xlSht.Cells[indexRow[i] + 1, 1] as Excel.Range;
                    name = cell.Value;
                    maxRow = indexRow[i + 1] - indexRow[i];

                }
                CreateTable(arrData, name, maxRow);

            }
           
            //xlApp.Visible = true; //отображаем Excel     
            xlWB.Close(false); //закрываем книгу, изменения не сохраняем
            xlApp.Quit(); //закрываем Excel

            //настройка DataGridView

        } 
        public static void CreateTable(dynamic worksheet, string namesave, int maxRow = 12)
        {
            Excel.Application ObjExcel = new Excel.Application();
            Excel.Workbook workbook = ObjExcel.Workbooks.Add(Type.Missing);
            Excel.Worksheet sheet = workbook.Worksheets.Add(Type.Missing);
            sheet.Range["A1:AI" + maxRow].Value = worksheet;

            workbook.SaveAs(namesave + ".xls");
            //ObjWorkBook = ObjExcel.Workbooks.Add(System.Reflection.Missing.Value);
            //ObjWorkBook.Worksheets.Add(  );
            //ObjExcel.Workbooks.Add(ObjWorkBook);
         
            ObjExcel.Quit();
        }
    }
}
