using System;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;

namespace ProcessData1018SCGLab1
{
    public static class DataHelpers
    {
        public static DataTable GetDataTableFromCsv(string path, bool isFirstRowHeader = true)
        {
            try
            {
                if (File.Exists(path))
                {
                    var header = isFirstRowHeader ? "Yes" : "No";
                    var pathOnly = Path.GetDirectoryName(path);
                    var fileName = Path.GetFileName(path);
                    var sql = $"SELECT * FROM [{fileName}]";
                    using (OleDbConnection connection = new OleDbConnection($"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={pathOnly};Extended Properties=\"Text;HDR={header}\""))
                    {
                        using (OleDbCommand command = new OleDbCommand(sql, connection))
                        {
                            using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                            {
                                var dataTable = new DataTable
                                {
                                    Locale = CultureInfo.CurrentCulture
                                };
                                adapter.Fill(dataTable);
                                return dataTable;
                            }
                        }
                    }
                }
                else
                {
                    Console.WriteLine(@"CSV load error; provided file does not exist");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            //default
            return null;
        }
    }
}