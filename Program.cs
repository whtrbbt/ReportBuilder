﻿using System;
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

            HouseReport(33445);
           
        }

        public static void HouseReport(int houseID)
        {
            DataTable report = new DataTable();
            DataTable houseTable = new DataTable();
            DataColumn column;
            DataRow row;

            //Задаем структуру таблицы houseTable----------------

            column = new DataColumn();
            column.ColumnName = "flat_num";
            houseTable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Double");
            column.ColumnName = "nach_val_start";
            houseTable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Double");
            column.ColumnName = "pay_val_start";
            houseTable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Double");
            column.ColumnName = "nach_val_now";
            houseTable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Double");
            column.ColumnName = "pay_val_now";
            houseTable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Double");
            column.ColumnName = "nach_val_end";
            houseTable.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Double");
            column.ColumnName = "pay_val_end";
            houseTable.Columns.Add(column);
            //------------------------------------------------------

            double cellValue = 0;
            double startCorrValue = 0;
            double nowCorrValue = 0;
            double endPeriodSaldo = 0;

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

            string queryString = $@"
                select house, flat_num, fls_full, resp_person,
                (select sum (val) from [ORACLE].[dbo].doc_nach where DOC_NACH.fls = fls_short and DOC_NACH.CREATED between '01.01.2014' and '31.12.2017' ) as nach_val_start,
                (select sum (val) from [ORACLE].[dbo].doc_pay where DOC_pay.fls = fls_short and (pay_date between '01.01.2014' and '31.12.2017') and (date_inp between '01.01.2014' and '31.12.2017') ) as pay_val_start,
                (select sum (val) from [ORACLE].[dbo].doc_correct where doc_correct.fls = fls_short and doc_correct.CREATED between '01.01.2014' and '31.12.2017' ) as cor_val_start,
                (select sum (val) from [ORACLE].[dbo].doc_nach where DOC_NACH.fls = fls_short and DOC_NACH.CREATED between '01.01.2018' and '31.03.2018' ) as nach_val_now,
                (select sum (val) from [ORACLE].[dbo].doc_pay where DOC_pay.fls = fls_short and  (date_inp between '01.01.2018' and '14.04.2018') ) as pay_val_now,
                (select sum (val) from [ORACLE].[dbo].doc_correct where doc_correct.fls = fls_short and doc_correct.CREATED between '01.01.2018' and '14.04.2018' ) as cor_val_now

                from 
                [ORACLE].[dbo].fls_view
                where
                fls_short in (select FLS_SHORT from [ORACLE].[dbo].fls_view where house ={houseID})
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

            da.Fill(report);
            //houseTable = report.Clone();
            


            foreach(DataRow active_row in report.Rows)
            {
                row = houseTable.NewRow();
                cellValue = 0;
                startCorrValue = 0;
                nowCorrValue = 0;
                endPeriodSaldo = 0;


                //Заполняем номер квартиры               
                row["flat_num"] = active_row["flat_num"];

                //Проверяем наличие корректировок
                if(!Convert.IsDBNull(active_row["cor_val_start"]))
                    startCorrValue = Convert.ToDouble(active_row["cor_val_start"]);
                else startCorrValue = 0;

                if(!Convert.IsDBNull(active_row["cor_val_now"]))
                    nowCorrValue = Convert.ToDouble(active_row["cor_val_now"]);
                else nowCorrValue = 0;


                //Обрабатываем данные на начало отчетного периода

                if((!Convert.IsDBNull(active_row["nach_val_start"])) & (!Convert.IsDBNull(active_row["pay_val_start"])))
                {                                 
					cellValue = Convert.ToDouble(active_row["nach_val_start"]) - Convert.ToDouble(active_row["pay_val_start"]) + startCorrValue;

                    if(cellValue > 0)
					{
                         row["nach_val_start"] = cellValue;
                         row["pay_val_start"] = 0;
					}
					else if(cellValue < 0)
					{
						cellValue = Math.Abs(cellValue);
                        row["nach_val_start"] = 0;
                        row["pay_val_start"] = cellValue;
                    }
					else
					{
                        row["nach_val_start"] = 0;
                        row["pay_val_start"] = 0;
                    }
                    
                }

                else if(!Convert.IsDBNull(active_row["nach_val_start"]))
                {
                    cellValue = Convert.ToDouble(active_row["nach_val_start"]) + startCorrValue;
                    
                    if(cellValue > 0)
                    {
                        row["nach_val_start"] = cellValue;
                        row["pay_val_start"] = 0;
                    }
                    else if(cellValue < 0)
                    {
                        cellValue = Math.Abs(cellValue);
                        row["nach_val_start"] = 0;
                        row["pay_val_start"] = cellValue;
                    }
                    else
                    {
                        row["nach_val_start"] = 0;
                        row["pay_val_start"] = 0;
                    }
                }
                
                else if(!Convert.IsDBNull(active_row["pay_val_start"]))
                {
                    cellValue = Convert.ToDouble(active_row["pay_val_start"]) - startCorrValue;
                    
                    if(cellValue > 0)
                    {
                        row["nach_val_start"] = cellValue;
                        row["pay_val_start"] = 0;
                    }
                    else if(cellValue < 0)
                    {
                        cellValue = Math.Abs(cellValue);
                        row["nach_val_start"] = 0;
                        row["pay_val_start"] = cellValue;
                    }
                    else
                    {
                        row["nach_val_start"] = 0;
                        row["pay_val_start"] = 0;
                    }
                }

                else
                {
                    cellValue = startCorrValue;

                    if(cellValue > 0)
                    {
                        row["nach_val_start"] = cellValue;
                        row["pay_val_start"] = 0;
                    }
                    else if(cellValue < 0)
                    {
                        cellValue = Math.Abs(cellValue);
                        row["nach_val_start"] = 0;
                        row["pay_val_start"] = cellValue;
                    }
                    else
                    {
                        row["nach_val_start"] = 0;
                        row["pay_val_start"] = 0;
                    }
                }


                //Обрабатываем данные за отчетный период
                if(!Convert.IsDBNull(active_row["nach_val_now"]))
                {
                    cellValue = Convert.ToDouble(active_row["nach_val_now"]);
                    row["nach_val_now"] = cellValue;
                }
                else
                {
                    row["nach_val_now"] = 0;
                }
                
                
                if(!Convert.IsDBNull(active_row["pay_val_now"]))
                {
                    cellValue = Convert.ToDouble(active_row["pay_val_now"]);
                    row["pay_val_now"] = cellValue;
                }
                else
                {
                    row["pay_val_now"] = 0;
                }


                //Обрабатываем данные на конец отчетного периода
                endPeriodSaldo = Convert.ToDouble(row["nach_val_start"]) - Convert.ToDouble(row["pay_val_start"]) + (Convert.ToDouble(row["nach_val_now"]) - Convert.ToDouble(row["pay_val_now"])
                    + nowCorrValue);
                
                if(endPeriodSaldo > 0)
                {
                    row["nach_val_end"] = endPeriodSaldo;
                    row["pay_val_end"] = 0;
                }
                else if(endPeriodSaldo < 0)
                {
                    endPeriodSaldo = Math.Abs(endPeriodSaldo);
                    row["nach_val_end"] = 0;
                    row["pay_val_end"] = endPeriodSaldo;
                }
                else
                {
                    row["nach_val_end"] = endPeriodSaldo;
                    row["pay_val_end"] = 0;
                }

                //Добавляем строку в таблицу для отчета если она не нулевая
                if((Convert.ToDouble(row["nach_val_start"]) + Convert.ToDouble(row["pay_val_start"]) + Convert.ToDouble(row["nach_val_now"]) + Convert.ToDouble(row["pay_val_now"])
                    + Convert.ToDouble(row["nach_val_end"]) + Convert.ToDouble(row["pay_val_end"])) != 0)
                    houseTable.Rows.Add(row);
                //иначе удаляем строку
                else row = null;
            }

            //Считаем итоги
            double totalStartNach = 0;  //итог по начислениям на начало периода
            double totalStartPay = 0;   //итог по платежам на начало периода
            double totalNowNach = 0;    //Итог по начислениям на момент отчета
            double totalNowPay = 0;     //Итог по платежам на момент отчета
            double totalEndNach = 0;    //Итог по начислениям на конец периода
            double totalEndPay = 0;     //Итог по платежам на конец периода
            
            foreach(DataRow total_row in houseTable.Rows)
            {
                totalStartNach += Convert.ToDouble(total_row["nach_val_start"]);
                totalStartPay += Convert.ToDouble(total_row["pay_val_start"]);
                totalNowNach += Convert.ToDouble(total_row["nach_val_now"]);
                totalNowPay += Convert.ToDouble(total_row["pay_val_now"]);
                totalEndNach += Convert.ToDouble(total_row["nach_val_end"]);
                totalEndPay += Convert.ToDouble(total_row["pay_val_end"]);
            }


            ///Выгружаем данные в Excel
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
            
            double startNach = 0;
            double startPay = 0;
            double nowNach = 0;
            double nowPay = 0;

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
                excelcells.Value2 = active_row["flat_num"];

                //Заполняем данные на начало отчетного периода
                excelcells = wsh.get_Range("B" + rowCounter, "B" + rowCounter);
                excelcells.Value2 = active_row["nach_val_start"];
                excelcells = wsh.get_Range("C" + rowCounter, "C" + rowCounter);
                excelcells.Value2 = active_row["pay_val_start"];
                excelcells = wsh.get_Range("D" + rowCounter, "D" + rowCounter);
                excelcells.Value2 = active_row["nach_val_now"];
                excelcells = wsh.get_Range("E" + rowCounter, "E" + rowCounter);
                excelcells.Value2 = active_row["pay_val_now"];
                excelcells = wsh.get_Range("G" + rowCounter, "G" + rowCounter);
                excelcells.Value2 = active_row["nach_val_end"];
                excelcells = wsh.get_Range("H" + rowCounter, "H" + rowCounter);
                excelcells.Value2 = active_row["pay_val_end"];

                rowCounter++;

            }

            //Форматируем итоговую таблицу
            Excel.Range tRange = wsh.get_Range(startCell, "I" + rowCounter);
            tRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            tRange.Borders.Weight = Excel.XlBorderWeight.xlThin;
            tRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            tRange.Font.Name = "Times New Roman";
            tRange.Font.Size = 9;
            tRange = wsh.get_Range("B" + ROWSHIFT, "I" + rowCounter);
            tRange.NumberFormat = "0.00";

            //Выводим итоги
            tRange = wsh.get_Range("A" + rowCounter, "A" + rowCounter);
            tRange.Value2 = "Итого: ";
            tRange = wsh.get_Range("B" + rowCounter, "B" + rowCounter);
            tRange.Value2 = totalStartNach;
            tRange = wsh.get_Range("C" + rowCounter, "C" + rowCounter);
            tRange.Value2 = totalStartPay;
            tRange = wsh.get_Range("D" + rowCounter, "D" + rowCounter);
            tRange.Value2 = totalNowNach;
            tRange = wsh.get_Range("E" + rowCounter, "E" + rowCounter);
            tRange.Value2 = totalNowPay;
            tRange = wsh.get_Range("G" + rowCounter, "G" + rowCounter);
            tRange.Value2 = totalEndNach;
            tRange = wsh.get_Range("H" + rowCounter, "H" + rowCounter);
            tRange.Value2 = totalEndPay;

            //Получаем адрес дома
            
            string fileName = "";
            DataTable houseAddr = GetHouseAddr(houseID);            
            foreach (DataRow dataRow in houseAddr.Rows)
            {
                fileName = Convert.ToString (dataRow["HOUSENUM"]);
                if(Convert.ToString(dataRow["HOUSECORP"])!="")
                    fileName += "_" + Convert.ToString(dataRow["HOUSECORP"]);
                if(Convert.ToString(dataRow["HOUSESUFIX"])!="")
                    fileName += "_" + Convert.ToString(dataRow["HOUSESUFIX"]);
                fileName += ".xlsx";
                break;
            }
            
            wb.SaveAs(fileName);
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
        }

        public static DataTable GetHouseAddr(int houseID)
        {
            DataTable house = new DataTable();
            
            DataTable fullAddr = new DataTable();
            DataColumn column;
            DataRow row;


            //Задаем структуру таблицы fullAddr----------------

            column = new DataColumn();
            column.ColumnName = "HOUSENUM";
            column.DataType = System.Type.GetType("System.String");
            fullAddr.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "HOUSECORP";
            fullAddr.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "HOUSESUFIX";
            fullAddr.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "AOGUID";
            fullAddr.Columns.Add(column);

            //column = new DataColumn();
            //column.DataType = System.Type.GetType("System.Double");
            //column.ColumnName = "pay_val_now";
            //fullAddr.Columns.Add(column);

            //column = new DataColumn();
            //column.DataType = System.Type.GetType("System.Double");
            //column.ColumnName = "nach_val_end";
            //fullAddr.Columns.Add(column);

            //column = new DataColumn();
            //column.DataType = System.Type.GetType("System.Double");
            //column.ColumnName = "pay_val_end";
            //fullAddr.Columns.Add(column);
            //------------------------------------------------------

            string houseAOGUID = "";
            string houseNum = "";
            string houseCorp = "";
            string houseSufix = "";
            
            
            
            SqlConnectionStringBuilder csbuilder = new SqlConnectionStringBuilder("");

            csbuilder["Server"] = @ConfigurationManager.AppSettings.Get("MSSQL_Server");
            csbuilder["UID"] = @ConfigurationManager.AppSettings.Get("UID");
            csbuilder["Password"] = @ConfigurationManager.AppSettings.Get("Password");
            csbuilder["Connect Timeout"] = 6000;
            csbuilder["integrated Security"] = true; //для коннекта с локальным экземпляром
            //csbuilder["Multisubnetfailover"] = "True";
            //csbuilder["Trusted_Connection"] = true;

            Console.WriteLine(csbuilder.ConnectionString);

            string queryString = $@"SELECT DISTINCT AOGUID, HOUSENUM, BUILDNUM, STRUCNUM FROM [ORACLE].dbo.FIAS_HOUSE_VIEW WHERE ADDR = {houseID}";

            SqlConnection conn = new SqlConnection(csbuilder.ConnectionString);
            SqlCommand cmd = new SqlCommand(queryString, conn);
            conn.Open();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(house);

            foreach (DataRow active_row in house.Rows)
            {
                houseAOGUID = Convert.ToString (active_row["AOGUID"]);
                houseNum = Convert.ToString(active_row["HOUSENUM"]);
                houseCorp = Convert.ToString(active_row["BUILDNUM"]);
                houseSufix = Convert.ToString(active_row["STRUCNUM"]);
                break;
            }

            row = fullAddr.NewRow();
            row["AOGUID"] = houseAOGUID;
            row["HOUSENUM"] = houseNum;
            row["HOUSECORP"] = houseCorp;
            row["HOUSESUFIX"] = houseSufix;
            fullAddr.Rows.Add(row);            
            conn.Close();
            da.Dispose();
            return fullAddr;
        }

        private static bool IsEven(int a)
        {
            return (a % 2) == 0;
        }

    }
}