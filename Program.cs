using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;

namespace NaicsStandard
{
    class Program
    {
        static DataTable getTableBySheet (string sheetName)
        {
            string fileName = @"..\\..\\..\\Naics.xlsx";
            string connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0}; Extended Properties=Excel 12.0;", fileName);

            OleDbConnection connection = new OleDbConnection(connectionString);
            connection.Open();
            DataTable data = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

            OleDbDataAdapter adapter = new OleDbDataAdapter("SELECT * from [" + sheetName + "$]", connectionString);

            DataSet ds = new DataSet();

            adapter.Fill(ds);

            connection.Close();

            return ds.Tables[0];
        }

        static NaicsStd convertRowToObject (DataRow row, DataTable footnotes)
        {
            string code = row[0].ToString();
            string revenue = row[2].ToString().Trim().Replace(",", "");
            if (revenue == "") revenue = "0";
            string employees = row[3].ToString().Trim().Replace(",", "");
            if (employees == "") employees = "0";
            string other = row[4].ToString().Length > 0 ? row[4].ToString().Substring(13) : "0";

            return new NaicsStd(code, revenue, employees, other);
        }

        static string getFootnote(DataRow row)
        {
            string untrimmed = row[1].ToString();
            int index = untrimmed.IndexOf("–");
            if (index < 0) return untrimmed;
            return untrimmed.Substring(index + 1).Trim();
        }

        static void generateNaicsStd ()
        {
            DataTable standards = getTableBySheet("standards");
            DataTable footnotes = getTableBySheet("footnotes");
            JObject naicsStd = new JObject();
            foreach (DataRow row in standards.Rows)
            {
                int x = 0;
                if (row[0].ToString().Length > 0 && int.TryParse(row[0].ToString(), out x))
                {
                    NaicsStd data = convertRowToObject(row, footnotes);
                    JObject code = new JObject();
                    code["revenue"] = data.revenue;
                    code["employees"] = data.employees;
                    code["footnote"] = data.other;
                    naicsStd[data.code] = code;
                }
            }
            System.IO.File.WriteAllText(@"C:\Users\nancy\source\repos\NaicsStandard\NaicsStandard\NaicsStd.json", JsonConvert.SerializeObject(naicsStd));
        }

        static void generateNaicsFootnote()
        {
            DataTable footnotes = getTableBySheet("footnotes");
            JObject naicsFootnotes = new JObject();
            string footnoteNum = "";
            for(int t = 0; t < footnotes.Rows.Count; t++)
            {
                DataRow row = footnotes.Rows[t];
                int x = 0;
                if (row[0].ToString().Length > 0 && int.TryParse(row[0].ToString(), out x))
                {
                    footnoteNum = x.ToString();
                    JArray footnoteArray = new JArray();
                    footnoteArray.Add(getFootnote(row));
                    naicsFootnotes[footnoteNum] = footnoteArray;
                }
                else if (row[1].ToString().Length > 0 && t != 0)
                {
                    JArray footnoteArray = naicsFootnotes[footnoteNum] as JArray;
                    if (footnoteArray != null)
                    {
                        footnoteArray.Add(getFootnote(row));
                        naicsFootnotes[footnoteNum] = footnoteArray;
                    }
                }
            }
            foreach(KeyValuePair<string, JToken> pair in naicsFootnotes)
            {
                JArray footnoteArr = pair.Value as JArray;
                if(footnoteArr.Count > 1 && ((string)footnoteArr[0]).StartsWith("NAICS"))
                {
                    footnoteArr.RemoveAt(0);
                    naicsFootnotes[pair.Key] = footnoteArr;
                }
            }
            System.IO.File.WriteAllText(@"C:\Users\nancy\source\repos\NaicsStandard\NaicsStandard\NaicsFootnote.json", JsonConvert.SerializeObject(naicsFootnotes));
        }

        static void Main(string[] args)
        {
            generateNaicsStd();
            generateNaicsFootnote();
        }
    }
}
