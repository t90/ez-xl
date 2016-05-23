using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ez_xl;

namespace SimpleExcelODS
{
    class Program
    {
        static void Main(string[] args)
        {
            var items = new[]
            {
                new ItemClass{Item1 = "line1", Item2 = 100},
                new ItemClass{Item1 = "line2", Item2 = 200},
                new ItemClass{Item1 = "line3", Item2 = 300},
            };

            if(File.Exists(@"result.ods"))File.Delete(@"result.ods");

            using (var excelWriter = new ExcelWriter(@"template.ods"))
            {
                excelWriter.Write(@"result.ods",new []
                {
                    new Tuple<string, IEnumerable<object>>("Sheet1",items),
                    new Tuple<string, IEnumerable<object>>("Sheet2",items.Take(2).ToList()),
                });
            }
        }
    }

    internal class ItemClass
    {
        public string Item1 { get; set; }
        public decimal Item2 { get; set; }
    }
}
