using Microsoft.VisualBasic.FileIO;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Text.RegularExpressions;


namespace CSVUtility
{
    public static class CSVUtility
    {
        public static void ToCSV(this DataTable dtDataTable, string strFilePath)
        {
            StreamWriter sw = new StreamWriter(strFilePath, false);
            //headers  

            //Паттерн для поиска разделителя в полях таблицы
            string pattern = ";+";
            
            for (int i = 0; i < dtDataTable.Columns.Count; i++)
            {
                sw.Write(dtDataTable.Columns[i]);
                if (i < dtDataTable.Columns.Count - 1)
                {
                    sw.Write(";");
                }
            }
            sw.Write(sw.NewLine);
            foreach(DataRow dr in dtDataTable.Rows)
            {
                for(int i = 0; i < dtDataTable.Columns.Count; i++)
                {
                    if(!Convert.IsDBNull(dr[i]))
                    {
                        string value = dr[i].ToString();
                        //if (value.Contains(';'))
                        //{
                        //value = String.Format("\\{0}\\", value);
                        value = Regex.Replace(value, @"\n+", " ");
                        value = Regex.Replace(value, pattern, ":");
                        sw.Write(value);
                        //}
                        //else
                        //{
                        //    sw.Write(dr[i].ToString());
                        //}
                    }
                    else if(Convert.IsDBNull(dr[i]))
                    {
                        sw.Write("NULL");
                    }
                    if(i < dtDataTable.Columns.Count - 1)
                    {
                        sw.Write(";");
                    }
                }
                sw.Write(sw.NewLine);
            }
            sw.Close();
        }

        public static DataTable GetDataTabletFromCSVFile(string csv_file_path)
        {
            Console.WriteLine(csv_file_path);

            //string[] lines = System.IO.File.ReadAllLines(csv_file_path);

            //// Display the file contents by using a foreach loop.
            //System.Console.WriteLine("Contents of PAY_DOC.CSV = ");
            //foreach (string line in lines)
            //{
            //    // Use a tab to indent each line of the file.
            //    Console.WriteLine("\t" + line);
            //}

            DataTable csvData = new DataTable();
            try
            {
                using (TextFieldParser csvReader = new TextFieldParser(csv_file_path))
                {
                    csvReader.SetDelimiters(new string[] { ";" });
                    csvReader.HasFieldsEnclosedInQuotes = false;
                    string[] colFields = csvReader.ReadFields();
                    Console.WriteLine("Количество столбцов: {0}", colFields.Length);
                    foreach (string column in colFields)
                    {
                        DataColumn datecolumn = new DataColumn(column);
                        //Console.WriteLine("Поле: {0}", column);
                        datecolumn.AllowDBNull = true;
                        csvData.Columns.Add(datecolumn);
                    }
                    while (!csvReader.EndOfData)
                    {
                        string[] fieldData = csvReader.ReadFields();
                        //Обработка null значений
                        for (int i = 0; i < fieldData.Length; i++)
                        {
                            if (fieldData[i] == "NULL")
                            {
                                fieldData[i] = null;
                            }
                        }
                        csvData.Rows.Add(fieldData);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: {0}", ex);
                return null;
            }
            return csvData;
        }


        public static void InsertDataIntoSQLServerUsingSQLBulkCopy(DataTable csvFileData, string tn, string cs)
        {
            using (SqlConnection dbConnection = new SqlConnection(cs))
            {
                dbConnection.Open();
                using (SqlBulkCopy s = new SqlBulkCopy(dbConnection))
                {
                    s.DestinationTableName = tn;
                    s.EnableStreaming = true;
                    s.BatchSize = 10000;
                    s.BulkCopyTimeout = 0;
                    s.NotifyAfter = 100;
                    s.SqlRowsCopied += delegate (object sender, SqlRowsCopiedEventArgs e)
                    {
                        Console.WriteLine(e.RowsCopied.ToString("#,##0") + " rows copied.");
                    };
                    foreach (var column in csvFileData.Columns)
                    {
                        s.ColumnMappings.Add(column.ToString(), column.ToString());

                        Console.WriteLine();
                    }
                    s.WriteToServer(csvFileData);
                }
                dbConnection.Close();
            }
        }


    }  
 
}

