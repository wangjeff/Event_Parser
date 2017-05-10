using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace EventConverterConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("GS Event Parser v1.5 - made by Jeff.Wang");
            try
            {
                FileOperation cFileOperation = new FileOperation();
                DataTable ControllerEventDataTable = new DataTable();
                DataTable SortEventDataTable = new DataTable();
                string FileName = args[0];
                //string FileName = "EventLogEnhancementRevised20160908.xls";
                string OutPutEventFileName = "OutputEventWord.doc";
                int OutPutEventFileNameIndex = 0;

                //Excel 的工作表名稱 (Excel左下角有的分頁名稱)
                // string SheetName = "ControllerEvent";
                //string EXEPath = System.Environment.CurrentDirectory + "\\TestSource.xlsx";
                string EXEPath = System.Environment.CurrentDirectory + "\\" + FileName;
                string EventFileOUTPath = System.Environment.CurrentDirectory + "\\" + OutPutEventFileName;

                /*確認檔案是否存在*/
                while (System.IO.File.Exists(EventFileOUTPath))
                {
                    EventFileOUTPath = System.Environment.CurrentDirectory + "\\" + OutPutEventFileNameIndex + "-" + OutPutEventFileName;
                    OutPutEventFileNameIndex++;
                }

                Console.WriteLine("Parsing ControllerEvent...");
                cFileOperation.ReadExcelFile(EXEPath, "Total events", ref ControllerEventDataTable);    //讀取controller event
                Console.WriteLine("Parsing ControllerEvent complete");

                Console.WriteLine("Sort Event...");
                cFileOperation.SortEventData(ref SortEventDataTable, ref ControllerEventDataTable); //合併排序event
                Console.WriteLine("Sort Event complete");

                Console.WriteLine("Write Event to Word...");
                cFileOperation.WriteWordFile(EventFileOUTPath, ref SortEventDataTable);    //將 event寫入

                Console.Write("Press any key to continue");
                Console.ReadLine();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine("please type valid form - Ex. EventConverterConsole.exe TestSource.xlsx");
                Console.ReadLine();
            }

        }
    }
}
