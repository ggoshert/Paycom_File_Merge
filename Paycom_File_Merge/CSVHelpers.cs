using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data;


namespace Paycom_File_Merge
{
    public static class CSVHelpers
    {
        public static void CombineCsvFiles(string sourceFolder, string destinationFile, string searchPattern, bool isMismatched)
        {
            // Specify wildcard search to match CSV files that will be combined
            string[] filePaths = Directory.GetFiles(sourceFolder, searchPattern);
            if (isMismatched)
                CombineMisMatchedCsvFiles(filePaths, destinationFile);
            else
                CombineCsvFiles(filePaths, destinationFile);
        }

        public static void CombineCsvFiles(string[] filePaths, string destinationFile)
        {
            StreamWriter fileDest = new StreamWriter(destinationFile, true);

            int i;
            for (i = 0; i < filePaths.Length; i++)
            {
                string file = filePaths[i];

                string[] lines = File.ReadAllLines(file);

                if (i > 0)
                {
                    lines = lines.Skip(1).ToArray(); // Skip header row for all but first file
                }

                foreach (string line in lines)
                {
                    fileDest.WriteLine(line);
                }
            }

            fileDest.Close();
        }

        public static void CombineMisMatchedCsvFiles(string[] filePaths, string destinationFile, char splitter = ',')
        {

            HashSet<string> combinedheaders = new HashSet<string>();
            int i;
            // aggregate headers
            for (i = 0; i < filePaths.Length; i++)
            {
                string file = filePaths[i];
                combinedheaders.UnionWith(File.ReadLines(file).First().Split(splitter));
            }
            var hdict = combinedheaders.ToDictionary(y => y, y => new List<object>());

            string[] combinedHeadersArray = combinedheaders.ToArray();
            for (i = 0; i < filePaths.Length; i++)
            {
                var fileheaders = File.ReadLines(filePaths[i]).First().Split(splitter);
                var notfileheaders = combinedheaders.Except(fileheaders);

                File.ReadLines(filePaths[i]).Skip(1).Select(line => line.Split(splitter)).ToList().ForEach(spline =>
                {
                    for (int j = 0; j < fileheaders.Length; j++)
                    {
                        hdict[fileheaders[j]].Add(spline[j]);
                    }
                    foreach (string header in notfileheaders)
                    {
                        hdict[header].Add(null);
                    }

                });
            }

            DataTable dt = hdict.ToDataTable();

            dt.ToCSV(destinationFile);
        }
    }

    public static class DataTableHelper
    {
        public static DataTable ToDataTable(this Dictionary<string, List<object>> dict)
        {
            DataTable dataTable = new DataTable();

            dataTable.Columns.AddRange(dict.Keys.Select(c => new DataColumn(c)).ToArray());

            for (int i = 0; i < dict.Values.Max(item => item.Count()); i++)
            {
                DataRow dataRow = dataTable.NewRow();

                foreach (var key in dict.Keys)
                {
                    if (dict[key].Count > i)
                        dataRow[key] = dict[key][i];
                }
                dataTable.Rows.Add(dataRow);
            }

            return dataTable;
        }

        public static void ToCSV(this DataTable dt, string destinationfile)
        {
            StringBuilder sb = new StringBuilder();

            IEnumerable<string> columnNames = dt.Columns.Cast<DataColumn>().
                                              Select(column => column.ColumnName);
            sb.AppendLine(string.Join(",", columnNames));

            foreach (DataRow row in dt.Rows)
            {
                IEnumerable<string> fields = row.ItemArray.Select(field => field.ToString());
                sb.AppendLine(string.Join(",", fields));
            }

            File.WriteAllText(destinationfile, sb.ToString());
        }

        public static void DataTableToCSV(this DataTable datatable, string destinationfile)
        {
            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < datatable.Columns.Count; i++)
            {
                sb.Append("\"" + datatable.Columns[i] + "\"");
                //if (i < datatable.Columns.Count - 1)
                //    sb.Append(seperator);
            }
            sb.AppendLine();
            foreach (DataRow dr in datatable.Rows)
            {
                for (int i = 0; i < datatable.Columns.Count; i++)
                {
                    sb.Append("\"" + dr[i].ToString() + "\"");

                    //if (i < datatable.Columns.Count - 1)
                    //    sb.Append(seperator);
                }
                sb.AppendLine();
            }
            //return sb.ToString();
            File.WriteAllText(destinationfile, sb.ToString());
        }

        public static void WriteCsv(this DataTable dt, string path)
        {
            using (var writer = new StreamWriter(path))
            {
                writer.WriteLine(string.Join(",", dt.Columns.Cast<DataColumn>().Select(dc => dc.ColumnName)));
                foreach (DataRow row in dt.Rows)
                {
                    writer.WriteLine(string.Join(",", row.ItemArray));
                }
            }
        }

        public static void ToCsv(this DataTable inDataTable, string destinationfile, bool inIncludeHeaders = true)
        {
            var builder = new StringBuilder();
            var columnNames = inDataTable.Columns.Cast<DataColumn>().Select(column => column.ColumnName);
            if (inIncludeHeaders)
                builder.AppendLine(string.Join(",", columnNames));
            foreach (DataRow row in inDataTable.Rows)
            {
                var fields = row.ItemArray.Select(field => field.ToString().WrapInQuotesIfContains(","));
                builder.AppendLine(string.Join(",", fields));
            }

            //return builder.ToString();
            File.WriteAllText(destinationfile, builder.ToString());
        }

        public static string WrapInQuotesIfContains(this string inString, string inSearchString)
        {
            if (inString.Contains(inSearchString))
                return "\"" + inString + "\"";
            return inString;
        }
    }
}
