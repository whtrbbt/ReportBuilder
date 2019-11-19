using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.IO;


namespace ExcelApp
{
    class Program
    {
        static void Main(string[] args)
        {

            DataTable table = new DataTable("Report");
            DataColumn column;
            DataRow row;

            string MSSQLtableName = @ConfigurationManager.AppSettings.Get("MSSQLtableName");


            SqlConnectionStringBuilder csbuilder =
                new SqlConnectionStringBuilder("");

            csbuilder["Server"] = @ConfigurationManager.AppSettings.Get("MSSQL_Server");
            csbuilder["UID"] = @ConfigurationManager.AppSettings.Get("UID");
            csbuilder["Password"] = @ConfigurationManager.AppSettings.Get("Password");
            csbuilder["Connect Timeout"] = 6000;
            csbuilder["integrated Security"] = true; //для коннекта с локальным экземпляром
            //csbuilder["Multisubnetfailover"] = "True";
            //csbuilder["Trusted_Connection"] = true;

            Console.WriteLine(csbuilder.ConnectionString);

            string queryString = @"
                select house, flat_num, fls_full, resp_person,
                (select sum (val) from [ORACLE].[dbo].doc_nach where DOC_NACH.fls = fls_short and DOC_NACH.CREATED between '01.01.2014' and '31.12.2017' ) as nach_val_start,
                (select sum (val) from [ORACLE].[dbo].doc_pay where DOC_pay.fls = fls_short and (pay_date between '01.01.2014' and '31.12.2017') and (date_inp between '01.01.2014' and '31.12.2017') ) as pay_val_start,
                (select sum (val) from [ORACLE].[dbo].doc_nach where DOC_NACH.fls = fls_short and DOC_NACH.CREATED between '01.01.2018' and '31.03.2018' ) as nach_val_now,
                (select sum (val) from [ORACLE].[dbo].doc_pay where DOC_pay.fls = fls_short and  (date_inp between '01.01.2018' and '16.04.2018') ) as pay_val_now

                from 
                [ORACLE].[dbo].fls_view
                where
                fls_short in (select FLS_SHORT from [ORACLE].[dbo].fls_view where house = 33445)

                --Сортируем квартры по возрастанию
                order by 
                    case IsNumeric(FLAT_NUM) 
                    when 1 then Replicate('0', 100 - Len(FLAT_NUM)) + FLAT_NUM
                    else FLAT_NUM
                end";

            SqlConnection conn = new SqlConnection(csbuilder.ConnectionString);
            SqlCommand cmd = new SqlCommand(queryString, conn);
            conn.Open();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            // this will query your database and return the result to your datatable

            DataTable houseTable = new DataTable();
            da.Fill(houseTable);

            //Объявляем приложение
            Excel.Application exc = new Microsoft.Office.Interop.Excel.Application();

            Excel.XlReferenceStyle RefStyle = exc.ReferenceStyle;

            Excel.Workbook wb = null;
            String TemplatePath = "house_report.xltx";
            try
            {
                wb = exc.Workbooks.Add(TemplatePath); // !!! 
            }
            catch(System.Exception ex)
            {
                throw new Exception("Не удалось загрузить шаблон для экспорта " + TemplatePath + "\n" + ex.Message);
            }
            Console.WriteLine("Шаблон найден, начинаю выгрузку.Это может занять несколько минут.");
            //Excel.Sheets excelsheets;

            //Выбираем третий лист
            Excel.Worksheet wsh = wb.Worksheets.get_Item(3) as Excel.Worksheet;

            Excel.Range excelcells;

            int rowCounter = 0;
            int ROWSHIFT = 7;
            double cellValue = 0;
            double startNach = 0;
            double startPay = 0;
            double nowNach = 0;
            double nowPay = 0;
            double endPeriodSaldo=0;

            rowCounter = +ROWSHIFT;
            

            string startCell = "A" + rowCounter;
            
            
            foreach(DataRow active_row in houseTable.Rows)
            {

                cellValue = 0;
                startNach = 0; //начисленно на начало отчетного периода
                startPay = 0;  //оплачено на начало отчетного периода
                nowNach = 0;   //начисленно в отчетном периоде
                nowPay = 0;    //оплачено в отчетном периоде     
                endPeriodSaldo = 0; //сальдо на конец периода

                //Заполняем таблицу Excel 
                
                //Заполняем номер квартиры
                excelcells = wsh.get_Range("A" + rowCounter, "A" + rowCounter);
                excelcells.Value2 = active_row[1];
                
                //Заполняем данные на начало отчетного периода
                excelcells = wsh.get_Range("B" + rowCounter, "B" + rowCounter);
                if((!Convert.IsDBNull(active_row[4])) & (!Convert.IsDBNull(active_row[5])))
                {
                    cellValue = Convert.ToDouble(active_row[4]) - Convert.ToDouble(active_row[5]);
                    if(cellValue > 0)
                    {
                        excelcells = wsh.get_Range("B" + rowCounter, "B" + rowCounter);
                        excelcells.Value2 = cellValue;
                        excelcells = wsh.get_Range("C" + rowCounter, "C" + rowCounter);
                        excelcells.Value2 = 0;
                        startNach = cellValue;
                        startPay = 0;
                    }
                    else if(cellValue < 0)
                    {
                        excelcells = wsh.get_Range("B" + rowCounter, "B" + rowCounter);
                        excelcells.Value2 = 0;
                        excelcells = wsh.get_Range("C" + rowCounter, "C" + rowCounter);
                        excelcells.Value2 = Math.Abs(cellValue);
                        startNach = 0;
                        startPay = Math.Abs(cellValue);
                    }
                    else
                    {
                        excelcells = wsh.get_Range("B" + rowCounter, "B" + rowCounter);
                        excelcells.Value2 = 0;
                        excelcells = wsh.get_Range("C" + rowCounter, "C" + rowCounter);
                        excelcells.Value2 = 0;
                        startNach = 0;
                        startPay = 0;
                    }
                }
                else if(!Convert.IsDBNull(active_row[4]))
                {
                    cellValue = Convert.ToDouble(active_row[4]);
                    excelcells = wsh.get_Range("B" + rowCounter, "B" + rowCounter);
                    excelcells.Value2 = cellValue;
                    excelcells = wsh.get_Range("C" + rowCounter, "C" + rowCounter);
                    excelcells.Value2 = 0;
                    startNach = cellValue;
                    startPay = 0;
                }
                else if (!Convert.IsDBNull(active_row[5]))
                {
                    cellValue = Convert.ToDouble(active_row[5]);
                    excelcells = wsh.get_Range("B" + rowCounter, "B" + rowCounter);
                    excelcells.Value2 = 0;
                    excelcells = wsh.get_Range("C" + rowCounter, "C" + rowCounter);
                    excelcells.Value2 = cellValue;
                    startNach = 0;
                    startPay = cellValue;
                }

                else
                {
                    excelcells = wsh.get_Range("B" + rowCounter, "B" + rowCounter);
                    excelcells.Value2 = 0;
                    excelcells = wsh.get_Range("C" + rowCounter, "C" + rowCounter);
                    excelcells.Value2 = 0;


                }
                
                
                //Заполняем данные за отчетный период
                if(!Convert.IsDBNull(active_row[6]))
                {
                    cellValue = Convert.ToDouble(active_row[6]);
                    excelcells = wsh.get_Range("D" + rowCounter, "D" + rowCounter);
                    excelcells.Value2 = cellValue;
                    nowNach = cellValue;
                }
                else
                {
                    excelcells = wsh.get_Range("D" + rowCounter, "D" + rowCounter);
                    excelcells.Value2 = 0;
                    nowNach = cellValue;
                }
                if(!Convert.IsDBNull(active_row[7]))
                {
                    cellValue = Convert.ToDouble(active_row[7]);
                    excelcells = wsh.get_Range("E" + rowCounter, "E" + rowCounter);
                    excelcells.Value2 = cellValue;
                    nowPay = cellValue;
                }
                else
                {
                    excelcells = wsh.get_Range("E" + rowCounter, "E" + rowCounter);
                    excelcells.Value2 = 0;
                    nowPay = 0;
                }


                //Заполняем данные на конец отчетного периода
                endPeriodSaldo = (startNach - startPay)+(nowNach-nowPay);         
                if(endPeriodSaldo > 0)
                {
                    excelcells = wsh.get_Range("G" + rowCounter, "G" + rowCounter);
                    excelcells.Value2 = endPeriodSaldo;
                    excelcells = wsh.get_Range("H" + rowCounter, "H" + rowCounter);
                    excelcells.Value2 = 0;
                }
                else if(endPeriodSaldo < 0)
                {
                    excelcells = wsh.get_Range("G" + rowCounter, "G" + rowCounter);
                    excelcells.Value2 = 0;
                    excelcells = wsh.get_Range("H" + rowCounter, "H" + rowCounter);
                    excelcells.Value2 = Math.Abs(endPeriodSaldo);
                }
                else
                {
                    excelcells = wsh.get_Range("G" + rowCounter, "G" + rowCounter);
                    excelcells.Value2 = 0;
                    excelcells = wsh.get_Range("H" + rowCounter, "H" + rowCounter);
                    excelcells.Value2 = 0;
                }

                

                rowCounter++;

            }
            
            //Форматируем итоговую таблицу
            Excel.Range tRange = wsh.get_Range(startCell, "I" + rowCounter);
            tRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            tRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
            tRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            tRange.Font.Name = "Times New Roman";
            tRange.Font.Size = 9;
            tRange = wsh.get_Range("B"+ ROWSHIFT, "I" + rowCounter);
            tRange.NumberFormat = "0.00";


            wb.SaveAs("1.xlsx");
            exc.Quit();
            conn.Close();
            da.Dispose();

            /*  Excel.Range c1 = (Excel.Range)wsh.Cells[3, 1];
              Excel.Range c2 = (Excel.Range)wsh.Cells[1 + grid.Rows.Count - 1, grid.Columns.Count];
              Excel.Range range = (Excel.Range)wsh.get_Range(c1, c2);
              range.Value2 = d;*/

            //Excel.Visible = true;
            /* //Отобразить Excel
             ex.Visible = true;
             //Количество листов в рабочей книге
             ex.SheetsInNewWorkbook = 2;
             //Добавить рабочую книгу
             Excel.Workbook workBook = ex.Workbooks.Add(Type.Missing);
             //Отключить отображение окон с сообщениями
             ex.DisplayAlerts = false;
             //Получаем первый лист документа (счет начинается с 1)
             Excel.Worksheet sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);
             //Название листа (вкладки снизу)
             sheet.Name = "Отчет за 13.12.2017";*/





            //Console.WriteLine("Hello World!");
        }

        private static void ShowTable(DataTable table)
        {
            foreach(DataColumn col in table.Columns)
            {
                Console.Write("{0,-14}", col.ColumnName);
            }
            Console.WriteLine();

            foreach(DataRow row in table.Rows)
            {
                foreach(DataColumn col in table.Columns)
                {
                    if(col.DataType.Equals(typeof(DateTime)))
                        Console.Write("{0,-14:d}", row[col]);
                    else if(col.DataType.Equals(typeof(Decimal)))
                        Console.Write("{0,-14:C}", row[col]);
                    else
                        Console.Write("{0,-14}", row[col]);
                }
                Console.WriteLine();
            }
        }

        private static bool IsEven(int a)
        {
            return (a % 2) == 0;
        }

    }
}
