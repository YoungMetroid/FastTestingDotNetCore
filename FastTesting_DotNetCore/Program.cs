using Npoi_Library_DotNetCore;
using System;
using System.Collections.Generic;
using UtilityLibrary.Loggers;

namespace FastTesting_DotNetCore
{
    class Program
    {
        private const string TestFolderPath = @"E:\TestFolderChida\";
        private const string DBString = "ConextionChida";

        static void Main(string[] args)
        {
            Logger logger = Logger.getInstance;

            logger.setLogPathandFile(TestFolderPath, "Error.log");

            List<List<object>> table = new List<List<object>>();
            try
            {
                string emailMani = "felipe.elizalde@toshibagcs.com";

                string key_Mani = "/qSs1RMHkyZXMsF1" + System.Environment.NewLine + "tyNJ3tk=";
                Console.WriteLine(key_Mani);
                NpoiExcelReadWrite npoiExcelReader = new NpoiExcelReadWrite(TestFolderPath + "Animals.xls");
                npoiExcelReader.setSheet(0);
                List<List<object>> newData = new List<List<object>>();

                List<object> row = new List<object>();
                row.AddRange(new List<object> { "Modify", "Modify", "Modify" });
                newData.AddRange(new List<List<object>> { row });

                row = new List<object>();
                row.AddRange(new List<object> { "Modify2", "Modify2", "Modify2" });
                newData.AddRange(new List<List<object>> { row });

                row = new List<object>();
                row.AddRange(new List<object> { "Modify3", "Modify3", "Modify3" });
                newData.AddRange(new List<List<object>> { row });

                row = new List<object>();
                row.AddRange(new List<object> { "Modify4", "Modify4", "Modify4" });
                newData.AddRange(new List<List<object>> { row });


                npoiExcelReader.WriteList_To_Excel(1,0,0,2,newData,1);

                npoiExcelReader.saveFile2(TestFolderPath + "Animals2.xls");
            }
            catch (Exception ex)
            {
                logger.logException(ex);
            }

            try
            {
                Console.WriteLine(table[0][0]);
            }
            catch (Exception ex)
            {

            }
            
        }
    }
}
