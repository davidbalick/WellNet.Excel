using System;
using System.Data;
using System.Diagnostics;
using System.Text;
using WellNet.Excel;

namespace TestConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {
            //new DataToExcelShowChanges().DataSetToExcel(TestData1(), @"H:\Projects\HCI_NVX_to_WN\TestData1.xlsx", false);
            new DataToExcelShowChanges().DataSetToExcel(TestData2(), @"H:\Projects\HCI_NVX_to_WN\TestData2.xlsx", false);
        }

        private static DataSet TestData1()
        {
            var random = new Random();
            var result = new DataSet("TestData1");
            var dataTable = new DataTable("TestTable1");
            result.Tables.Add(dataTable);
            dataTable.Columns.Add("String column", typeof(string));
            dataTable.Columns.Add("Long string column", typeof(string));
            dataTable.Columns.Add("Integer column", typeof(int));
            dataTable.Columns.Add("Date column", typeof(DateTime));
            dataTable.Columns.Add("Currency column", typeof(double));
            for (var i = 0; i < random.Next(10,1000); i++)
                AddTestData1Row(dataTable, random);
            return result;
        }

        private static DataSet TestData2()
        {
            var result = new DataSet("TestData1");
            var dataTable = new DataTable("TestTable1");
            result.Tables.Add(dataTable);
            dataTable.Columns.Add("Id", typeof(string));
            dataTable.Columns.Add("NVX stuff", typeof(string));
            dataTable.Columns.Add("WN stuff", typeof(string));
            dataTable.Columns.Add("NVX Age", typeof(int));
            dataTable.Columns.Add("WN Age", typeof(int));
            var newRow = dataTable.NewRow();
            newRow[0] = "Person 1";
            newRow[1] = "abc";
            newRow[2] = "ABC";
            dataTable.Rows.Add(newRow);
            newRow = dataTable.NewRow();
            newRow[0] = "Person 2";
            newRow[1] = "def";
            newRow[2] = "def";
            newRow[3] = 47;
            newRow[4] = 46;
            dataTable.Rows.Add(newRow);
            newRow = dataTable.NewRow();
            newRow[0] = "Person 3";
            newRow[1] = "ghi";
            newRow[2] = "jkl";
            newRow[3] = 32;
            newRow[4] = 32;
            dataTable.Rows.Add(newRow);
            newRow = dataTable.NewRow();
            newRow[0] = "Person 4";
            newRow[2] = "jkl";
            newRow[3] = 64;
            dataTable.Rows.Add(newRow);
            return result;
        }

        private static void AddTestData1Row(DataTable dataTable, Random random)
        {
            var newRow = dataTable.NewRow();
            newRow[0] = LoremIpsum(1, 4, random);
            newRow[1] = LoremIpsum(5, 50, random);
            newRow[2] = random.Next();
            newRow[3] = DateTime.Now.AddDays(random.Next(1000));
            newRow[4] = random.NextDouble() * 100 * (random.Next(0,19) < 5 ? -1 : 1);
            dataTable.Rows.Add(newRow);
        }

        static string LoremIpsum(int minWords, int maxWords, Random rand)//, int minSentences, int maxSentences, int numParagraphs)
        {

            var words = new[]{"lorem", "ipsum", "dolor", "sit", "amet", "consectetuer",
                "adipiscing", "elit", "sed", "diam", "nonummy", "nibh", "euismod",
                "tincidunt", "ut", "laoreet", "dolore", "magna", "aliquam", "erat"};

            //int numSentences = rand.Next(maxSentences - minSentences)
            //    + minSentences + 1;
            int numWords = rand.Next(maxWords - minWords) + minWords + 1;

            StringBuilder result = new StringBuilder();

            //for (int p = 0; p < numParagraphs; p++)
            //{
            //    result.Append("<p>");
            //    for (int s = 0; s < numSentences; s++)
            //    {
                    for (int w = 0; w < numWords; w++)
                    {
                        if (w > 0) { result.Append(" "); }
                        result.Append(words[rand.Next(words.Length)]);
                    }
            //        result.Append(". ");
            //    }
            //    result.Append("</p>");
            //}

            return result.ToString();
        }
    }
}
