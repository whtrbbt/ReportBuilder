using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using CSVUtility;


namespace ExcelApp
{
    class Program
    {
        static void Main(string[] args)
        {

            // Определяем режим работы: пакетный или нет
            if (ConfigurationManager.AppSettings.Get("BATCH_MODE") == "1")
            {
                BatchMode(@ConfigurationManager.AppSettings.Get("BATCH_MAP"));
            }
            else
            {
                //Проверяем признак наличия улиц и выбираем нужный метод
                if (ConfigurationManager.AppSettings.Get("FORCE_NO_STREET") == "0")
                {
                    if (CityStreetHaveName(@ConfigurationManager.AppSettings.Get("CITYGUID")))
                        CityReport(@ConfigurationManager.AppSettings.Get("CITYGUID"), @ConfigurationManager.AppSettings.Get("CITY"));
                    else
                        StreetReport(@ConfigurationManager.AppSettings.Get("CITYGUID"), @ConfigurationManager.AppSettings.Get("CITY"), "", "нет улицы"); //для НП без улиц
                }
                else
                    StreetReport(@ConfigurationManager.AppSettings.Get("CITYGUID"), @ConfigurationManager.AppSettings.Get("CITY"), "", "нет улицы"); //для НП без улиц
                                                                                                      //CityReport("30fa6bcc-608e-40cb-b22a-b202967ff2a6");
                                                                                                      //StreetReport("003eb85c-27a7-41fc-b0c1-ffefc4b98755");
                                                                                                      //HouseReport(1068514, "1","1","1","1","1");
            }
        }


        public static void BatchMode (string batchFileName)
        {
            DataTable reportMap = new DataTable();
            reportMap = CSVUtility.CSVUtility.GetDataTabletFromCSVFile(batchFileName);
            foreach (DataRow dr in reportMap.AsEnumerable())
            {
                //Проверяем признак наличия улиц и выбираем нужный метод
                if (dr["STREET"].ToString() == "1")
                {
                    if (CityStreetHaveName(dr["AOGUID"].ToString()))
                        CityReport(dr["AOGUID"].ToString(), dr["NP"].ToString());
                    else
                        StreetReport(dr["AOGUID"].ToString(), dr["NP"].ToString(), "", "нет улицы"); //для НП без улиц
                }
                else
                    StreetReport(dr["AOGUID"].ToString(), dr["NP"].ToString(), "", "нет улицы");
            }
        }

        public static void HouseReport(int houseID, string cityName, string streetType, string streetName, string houseNum, string houseCorp, string houseSufix)
        //Формирует отчет по дому
        {
            DataTable report = new DataTable();
            DataTable houseTable = new DataTable();
            DataColumn column;
            DataRow row;
            string path = @ConfigurationManager.AppSettings.Get("PATH");
            string TEMPL_PATH = @ConfigurationManager.AppSettings.Get("TEMPL_PATH");

            #region Задаем структуру таблицы houseTable

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
            #endregion

            double cellValue = 0;
            double startCorrValue = 0;
            double nowCorrValue = 0;
            double endPeriodSaldo = 0;
            string reportPeriod = @ConfigurationManager.AppSettings.Get("ReportPeriod");
            double reportString = 0;

            DateTime START_REPORT_PERIOD = Convert.ToDateTime(@ConfigurationManager.AppSettings.Get("StartReportPeriod"));
            DateTime END_REPORT_PERIOD = Convert.ToDateTime(@ConfigurationManager.AppSettings.Get("EndReportPeriod"));
            DateTime PREVIOS_REPORT_PERIOD_END = START_REPORT_PERIOD.AddDays(-1);
            DateTime PAY_DAY = Convert.ToDateTime(@ConfigurationManager.AppSettings.Get("PayDay"));
            string startReportPeriod = @START_REPORT_PERIOD.ToString("dd.MM.yyyy");
            string endReportPeriod = @END_REPORT_PERIOD.ToString("dd.MM.yyyy");
            string previosReportPeriodEnd = @PREVIOS_REPORT_PERIOD_END.ToString("dd.MM.yyyy");
            string payDay = @PAY_DAY.ToString("dd.MM.yyyy");
            string startReportYear = "01.01." + @END_REPORT_PERIOD.ToString("yyyy");



            SqlConnectionStringBuilder csbuilder =
                new SqlConnectionStringBuilder("");

            csbuilder["Server"] = @ConfigurationManager.AppSettings.Get("MSSQL_Server");
            csbuilder["UID"] = @ConfigurationManager.AppSettings.Get("UID");
            csbuilder["Password"] = @ConfigurationManager.AppSettings.Get("Password");
            csbuilder["Connect Timeout"] = 6000;
            csbuilder["integrated Security"] = true; //для коннекта с локальным экземпляром

                        
            string queryString = $@"
                select house, flat_num, fls_full, resp_person,
                (select sum (val) from [ORACLE].[dbo].doc_nach where DOC_NACH.fls = fls_short and DOC_NACH.CREATED between '01.01.2014' and '{previosReportPeriodEnd}' ) as nach_val_start,
                (select sum (val) from [ORACLE].[dbo].doc_pay where DOC_pay.fls = fls_short and (pay_date between '01.01.2014' and '{previosReportPeriodEnd}') and (date_inp between '01.01.2014' and '{previosReportPeriodEnd}') ) as pay_val_start,
                (select sum (val) from [ORACLE].[dbo].doc_correct where doc_correct.fls = fls_short and doc_correct.CREATED between '01.01.2014' and '{previosReportPeriodEnd}' ) as cor_val_start,
                (select sum (val) from [ORACLE].[dbo].doc_nach where DOC_NACH.fls = fls_short and DOC_NACH.CREATED between '{startReportYear}' and
                '{endReportPeriod}' ) as nach_val_now,
                (select sum (val) from [ORACLE].[dbo].doc_pay where DOC_pay.fls = fls_short and  (date_inp between '{startReportYear}' and '{payDay}') ) as pay_val_now,
                (select sum (val) from [ORACLE].[dbo].doc_correct where doc_correct.fls = fls_short and doc_correct.CREATED between '{startReportPeriod}' and '{payDay}' ) as cor_val_now,
		        (select sum (val) from [ORACLE].[dbo].doc_nach where DOC_NACH.fls = fls_short and DOC_NACH.CREATED between '01.01.2014' and '{endReportPeriod}' ) as nach_val_end,
		        (select sum (val) from [ORACLE].[dbo].doc_pay where DOC_pay.fls = fls_short and  (date_inp between '01.01.2014' and '{payDay}') ) as pay_val_end,
                (select sum (val) from [ORACLE].[dbo].doc_pay where DOC_pay.fls = fls_short and  (date_inp between '{startReportPeriod}' and '{payDay}') ) as pay_val_period


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
            
            da.Fill(report);

            double totalPayNow = 0; //Итог по платежам в отчетном периоде
            double totalFoundNow = 0; //Итог по всем поступлениям за отчетный период
            double totalPay = 0; // Итог по всем платежам

            #region готовим данные для отчета
            foreach (DataRow active_row in report.Rows)
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

                //Собираем итог по платежам за отчетный период
                if(!Convert.IsDBNull(active_row["pay_val_period"]))
                {
                    cellValue = Convert.ToDouble(active_row["pay_val_period"]);                   
                    totalPayNow += cellValue;
                }


                //Обрабатываем данные на конец отчетного периода
                if((!Convert.IsDBNull(active_row["nach_val_end"])) & (!Convert.IsDBNull(active_row["pay_val_end"])))
                {
                    endPeriodSaldo = Convert.ToDouble(active_row["nach_val_end"]) - Convert.ToDouble(active_row["pay_val_end"]) + nowCorrValue;
                    totalPay += Convert.ToDouble(active_row["pay_val_end"]);
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
                }

                else if(!Convert.IsDBNull(active_row["nach_val_end"]))
                {
                    endPeriodSaldo = Convert.ToDouble(active_row["nach_val_end"]) + nowCorrValue;

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
                    else //endPeriodSaldo = 0
                    {
                        row["nach_val_end"] = endPeriodSaldo;
                        row["pay_val_end"] = 0;
                    }
                }

                else if(!Convert.IsDBNull(active_row["pay_val_end"]))
                {
                    endPeriodSaldo = Convert.ToDouble(active_row["pay_val_end"]) - nowCorrValue;
                    totalPay += Convert.ToDouble(active_row["pay_val_end"]);

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
                    else //endPeriodSaldo = 0
                    {
                        row["nach_val_end"] = endPeriodSaldo;
                        row["pay_val_end"] = 0;
                    }
                }

                else
                {
                    endPeriodSaldo = nowCorrValue;

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
                    else //endPeriodSaldo = 0
                    {
                        row["nach_val_end"] = endPeriodSaldo;
                        row["pay_val_end"] = 0;
                    }
                }

                //Расчитываем контрольную сумму по строке
                
                reportString = (Convert.ToDouble(row["nach_val_start"]) + Convert.ToDouble(row["pay_val_start"]) + Convert.ToDouble(row["nach_val_now"]) + Convert.ToDouble(row["pay_val_now"])
                    + Convert.ToDouble(row["nach_val_end"]) + Convert.ToDouble(row["pay_val_end"]));
                //Добавляем строку в таблицу для отчета если она не нулевая                
                if (reportString != 0)
                    houseTable.Rows.Add(row);
                //иначе удаляем строку
                else row = null; //reportString = 0
            }
            #endregion

            #region Считаем итоги
            double startBalance = 0;    //остаток средств на начало отчетного периода по дому
            double endBalance = 0;      //остаток средств на конец отчетного периода по дому
            double repairVolume = GetHouseRepairVolume(houseID);
                                        //Объем выполненных работ по капитальному ремонту дома в рублях

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
            

            #endregion

            #region Выгружаем данные в Excel
            //Объявляем приложение
            Excel.Application exc = new Microsoft.Office.Interop.Excel.Application();

            Excel.XlReferenceStyle RefStyle = exc.ReferenceStyle;

            Excel.Workbook wb = null;          
            
            try
            {
                wb = exc.Workbooks.Add(TEMPL_PATH); // !!! 
            }
            catch(System.Exception ex)
            {
                throw new Exception("Не удалось загрузить шаблон для экспорта " + TEMPL_PATH + "\n" + ex.Message);
            }

            #region заполняем первый лист

            //Заполняем реквизиты отчета

            Excel.Worksheet wsh1 = wb.Worksheets.get_Item(1) as Excel.Worksheet;

            //Выводим адрес дома
            string houseAddr = "";
            houseAddr = @ConfigurationManager.AppSettings.Get("OBL") + ", " + cityName +", "+
                streetType + " " + streetName + ", д." + houseNum + " " + houseCorp + " "+ houseSufix;
            Excel.Range titulRange = wsh1.get_Range("C7", "C7");
            titulRange.Value2 = houseAddr;

            //Выводим период отчета
            titulRange = wsh1.get_Range("C8", "C8");
            titulRange.Value2 = reportPeriod;

            //Выводим итоги по платежам на первый лист
            totalFoundNow = totalPayNow;
            startBalance = totalPay - repairVolume;
            endBalance = startBalance + totalFoundNow;
            Excel.Range totalRange;            
            totalRange = wsh1.get_Range("E16", "E16"); // 4.Поступило в отчетном периоде взносов (за счет минимального взноса) 
            totalRange.Value2 = Math.Round(totalPayNow / 1000, 2, MidpointRounding.AwayFromZero);
            totalRange = wsh1.get_Range("E18", "E18"); // 4.Поступило в отчетном периоде взносов (итого)
            totalRange.Value2 = Math.Round(totalPayNow / 1000, 2, MidpointRounding.AwayFromZero);
            totalRange = wsh1.get_Range("D16", "D16"); // 3.Поступило в отчетном периоде всего (за счет минимального взноса)
            totalRange.Value2 = Math.Round(totalFoundNow / 1000, 2, MidpointRounding.AwayFromZero);
            totalRange = wsh1.get_Range("D18", "D18"); // 3.Поступило в отчетном периоде всего (итого)
            totalRange.Value2 = Math.Round(totalFoundNow / 1000, 2, MidpointRounding.AwayFromZero);
            totalRange = wsh1.get_Range("C16", "C16"); // 2.Остаток средств на начало отчетного периода (за счет минимального взноса)
            totalRange.Value2 = Math.Round(startBalance / 1000, 2, MidpointRounding.AwayFromZero);
            totalRange = wsh1.get_Range("C18", "C18"); // 2.Остаток средств на начало отчетного периода (итого)
            totalRange.Value2 = Math.Round(startBalance / 1000, 2, MidpointRounding.AwayFromZero);
            totalRange = wsh1.get_Range("J16", "J16"); // 9.Остаток средств на конец отчетного периода (за счет минимального взноса)
            totalRange.Value2 = Math.Round(endBalance / 1000, 2, MidpointRounding.AwayFromZero);
            totalRange = wsh1.get_Range("J18", "J18"); // 9.Остаток средств на конец отчетного периода (итого)
            totalRange.Value2 = Math.Round(endBalance / 1000, 2, MidpointRounding.AwayFromZero);
            #endregion

            #region заполняем второй лист
            
            Excel.Worksheet wsh2 = wb.Worksheets.get_Item(2) as Excel.Worksheet;
            Excel.Range repairRange;

            string workName = "Кап. ремонт";

            repairRange = wsh2.get_Range("B7", "B7"); // 1.Виды работ и услуг по капитальному ремонту
            repairRange.Value2 = workName;
            repairRange = wsh2.get_Range("C7", "C7"); // 2.Стоимость работ и услуг по капитальному ремонту (кап. ремонт)
            repairRange.Value2 = Math.Round(repairVolume / 1000, 2, MidpointRounding.AwayFromZero);
            repairRange = wsh2.get_Range("C8", "C8"); // 2.Стоимость работ и услуг по капитальному ремонту (итого)
            repairRange.Value2 = Math.Round(repairVolume / 1000, 2, MidpointRounding.AwayFromZero);
            repairRange = wsh2.get_Range("D7", "D7"); // 3.Размер средств направленных на капитальный ремонт (кап. ремонт)
            repairRange.Value2 = Math.Round(repairVolume / 1000, 2, MidpointRounding.AwayFromZero);
            repairRange = wsh2.get_Range("D8", "D8"); // 3.Размер средств направленных на капитальный ремонт (итого)
            repairRange.Value2 = Math.Round(repairVolume / 1000, 2, MidpointRounding.AwayFromZero);

            #endregion

            #region заполняем третий лист
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
            tRange = wsh.get_Range("A" + rowCounter, "H" + rowCounter);
            tRange.Font.Bold = true;

            #endregion

            //Получаем адрес дома для формирования имени файла
            string fileName = path + cityName + "\\"+streetType + streetName;

            if(!Directory.Exists(fileName))
                Directory.CreateDirectory(fileName);

            string houseName = houseNum;
            if(houseSufix != "")
                houseName = houseName + "_" + houseSufix;
            if(houseCorp != "")                
                houseName = houseName + "к" + houseCorp;


            fileName += "\\" + RemoveInvalidChars(houseName) + ".xlsx";
            Console.WriteLine(fileName);

            wb.SaveAs(fileName);
            exc.Quit();
            #endregion
            conn.Close();
            da.Dispose();

        }

        public static DataTable GetHouseAddr(int houseID)
        //Получаем адрес дома
        {
            DataTable house = new DataTable();
            
            DataTable fullAddr = new DataTable();
            DataColumn column;
            DataRow row;


            #region Задаем структуру таблицы fullAddr

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

            #endregion

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

        public static double GetHouseRepairVolume (int houseID)
        //Получаем объем выполненых работ по кап. ремонту дома в рублях
        {
            double repairVolume = 0;
            DataTable house = new DataTable();

            SqlConnectionStringBuilder csbuilder = new SqlConnectionStringBuilder("");

            csbuilder["Server"] = @ConfigurationManager.AppSettings.Get("MSSQL_Server");
            csbuilder["UID"] = @ConfigurationManager.AppSettings.Get("UID");
            csbuilder["Password"] = @ConfigurationManager.AppSettings.Get("Password");
            csbuilder["Connect Timeout"] = 6000;
            csbuilder["integrated Security"] = true; //для коннекта с локальным экземпляром

            string queryString = $@"SELECT 
                                    [ADDR],SUM([VOL]) VOL,[ADDR_STR]
                                    FROM [ORACLE].[dbo].[REMONT_VOLUME]
                                    WHERE ADDR = {houseID}
                                    GROUP BY ADDR, ADDR_STR";

            SqlConnection conn = new SqlConnection(csbuilder.ConnectionString);
            SqlCommand cmd = new SqlCommand(queryString, conn);
            conn.Open();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(house);

            if (house.Rows.Count > 0)
                foreach (DataRow dr in house.AsEnumerable())
                    repairVolume = Convert.ToDouble(dr["VOL"]);
            else //house.Rows.Count <= 0
                repairVolume = 0;
            
            conn.Close();
            da.Dispose();
            return repairVolume;
        }


        public static void StreetReport (string aoGUID, string cityName, string streetType, string streetName)
        //Формирует отчет по всем домам на улице
        {
            DataTable houses = new DataTable();            
           
            DataColumn column;
            DataRow row;          
            
            SqlConnectionStringBuilder csbuilder = new SqlConnectionStringBuilder("");

            csbuilder["Server"] = @ConfigurationManager.AppSettings.Get("MSSQL_Server");
            csbuilder["UID"] = @ConfigurationManager.AppSettings.Get("UID");
            csbuilder["Password"] = @ConfigurationManager.AppSettings.Get("Password");
            csbuilder["Connect Timeout"] = 6000;
            csbuilder["integrated Security"] = true; //для коннекта с локальным экземпляром      

            string queryString = $@"SELECT ADDR, HOUSENUM, BUILDNUM, STRUCNUM  FROM [ORACLE].dbo.FIAS_HOUSE_VIEW WHERE AOGUID = '{aoGUID}'";
            try
            {
                SqlConnection conn = new SqlConnection(csbuilder.ConnectionString);
                SqlCommand cmd = new SqlCommand(queryString, conn);
                conn.Open();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(houses);

                int houseID = 0;
                string houseCorp = "";
                string houseSufix = "";

                Console.WriteLine(streetName);

                foreach (DataRow active_row in houses.Rows)
                {
                    houseCorp = "";
                    houseSufix = "";
                    if (!Convert.IsDBNull(active_row["BUILDNUM"]))
                        houseCorp = @Convert.ToString(active_row["BUILDNUM"]);
                    if (!Convert.IsDBNull(active_row["STRUCNUM"]))
                        houseSufix = @Convert.ToString(active_row["STRUCNUM"]);

                    HouseReport(Convert.ToInt32(active_row["ADDR"]), cityName,streetType, streetName, Convert.ToString(active_row["HOUSENUM"]), houseCorp, houseSufix);
                }
                conn.Close();

            }

            catch (System.Exception ex)
            {
                throw new Exception("Ошибка при получении данных из БД: " + ex.Message);
            }




        }

        public static void CityReport(string cityGUID, string cityName)
        //Формирует отчет по всем домам в городе
        {
            DataTable streets = new DataTable();
            DataColumn column;
            DataRow row;

            Console.WriteLine(cityName);

            SqlConnectionStringBuilder csbuilder = new SqlConnectionStringBuilder("");

            csbuilder["Server"] = @ConfigurationManager.AppSettings.Get("MSSQL_Server");
            csbuilder["UID"] = @ConfigurationManager.AppSettings.Get("UID");
            csbuilder["Password"] = @ConfigurationManager.AppSettings.Get("Password");
            csbuilder["Connect Timeout"] = 6000;
            csbuilder["integrated Security"] = true; //для коннекта с локальным экземпляром     

            string queryString = $@"SELECT distinct AOGUID, OFFNAME, SHORTNAME, AOLEVEL
                                    FROM [ORACLE].[dbo].[FIAS_ADDROBJ_LOAD]
                                    where PARENTGUID = '{cityGUID}' and AOLEVEL = 7 and ACTSTATUS = 1
                                    order by aoguid";

            SqlConnection conn = new SqlConnection(csbuilder.ConnectionString);
            SqlCommand cmd = new SqlCommand(queryString, conn);
            conn.Open();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(streets);



            foreach(DataRow active_row in streets.Rows)
            {
                StreetReport(Convert.ToString(active_row["AOGUID"]), cityName, Convert.ToString(active_row["SHORTNAME"]), Convert.ToString(active_row["OFFNAME"]));                
            }

            conn.Close();
            da.Dispose();
        }

        public static String RemoveInvalidChars(String file_name)
        //Убирает недопустимые символы из имени файла
        
        {
            foreach(Char invalid_char in Path.GetInvalidFileNameChars())
            {
                file_name = file_name.Replace(oldValue: invalid_char.ToString(), newValue: "_");
            }
            return file_name;
        }

        public static bool CityStreetHaveName (string cityGUID)
        //Проверяет наличе улиц в нас. пункте
        {
            DataTable houses = new DataTable();
            SqlConnectionStringBuilder csbuilder = new SqlConnectionStringBuilder("");
            csbuilder["Server"] = @ConfigurationManager.AppSettings.Get("MSSQL_Server");
            csbuilder["UID"] = @ConfigurationManager.AppSettings.Get("UID");
            csbuilder["Password"] = @ConfigurationManager.AppSettings.Get("Password");
            csbuilder["Connect Timeout"] = 6000;
            csbuilder["integrated Security"] = true; //для коннекта с локальным экземпляром      

            string queryString = $@"SELECT distinct AOGUID
                                    FROM [ORACLE].[dbo].[FIAS_ADDROBJ_LOAD]
                                    where PARENTGUID = '{cityGUID}' and AOLEVEL = 7 and ACTSTATUS = 1
                                    order by aoguid";

            SqlConnection conn = new SqlConnection(csbuilder.ConnectionString);
            SqlCommand cmd = new SqlCommand(queryString, conn);
            conn.Open();
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(houses);

            if(houses.Rows.Count > 0)
            {
                conn.Close();
                da.Dispose();
                return true;
            }
            else
            {
                conn.Close();
                da.Dispose();
                return false;
            }
         }
        
        private static bool IsEven(int a)
        {
            return (a % 2) == 0;
        }

    }
}
