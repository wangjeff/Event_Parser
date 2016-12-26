using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.OleDb;
using System.Data;
using Microsoft.Office.Interop.Word;

namespace EventConverterConsole
{
    class FileOperation
    {
        //2.提供者名稱  Microsoft.Jet.OLEDB.4.0適用於2003以前版本，Microsoft.ACE.OLEDB.12.0 適用於2007以後的版本處理 xlsx 檔案
        private const string ProviderName = "Microsoft.ACE.OLEDB.12.0;";
        //3.Excel版本，Excel 8.0 針對Excel2000及以上版本，Excel5.0 針對Excel97。
        private const string ExtendedString = "'Excel 8.0;";
        //4.第一行是否為標題
        private const string Hdr = "YES;";
        //5.IMEX=1 通知驅動程序始終將「互混」數據列作為文本讀取
        private const string IMEX = "0';";

        public void MergeSortEventDataTable(ref System.Data.DataTable MergeSortDataTable, ref System.Data.DataTable ControllerDataTable, ref System.Data.DataTable DataServiceDataTable)
        {
            try
            {
                System.Data.DataRow[] InformationDataRow;
                System.Data.DataRow[] WarningDataRow;
                System.Data.DataRow[] ErrorRow;
                System.Data.DataRow[] CriticalErrorRow;
                System.Data.DataTable MergeDataTable = new System.Data.DataTable();
                System.Data.DataTable InformationDataTable = new System.Data.DataTable();
                System.Data.DataTable WarningDataTable = new System.Data.DataTable();
                System.Data.DataTable ErrorDataTable = new System.Data.DataTable();
                System.Data.DataTable CriticalErrorDataTable = new System.Data.DataTable();

                //Merge dataservice and controller event
                MergeDataTable = ControllerDataTable.Copy();
                MergeDataTable.Merge(DataServiceDataTable, true);
                
                //依序query sevirity完成排序
                InformationDataRow = MergeDataTable.Select("Severity='Information'");
                WarningDataRow = MergeDataTable.Select("Severity='Warning'");
                ErrorRow = MergeDataTable.Select("Severity='Error'");
                CriticalErrorRow = MergeDataTable.Select("Severity='Critical Error'");

                InformationDataTable = MergeDataTable.Clone();
                WarningDataTable = MergeDataTable.Clone();
                ErrorDataTable = MergeDataTable.Clone();
                CriticalErrorDataTable = MergeDataTable.Clone();

                for (int i = 0; i < InformationDataRow.Length; i++)
                {
                    InformationDataTable.ImportRow(InformationDataRow[i]);
                }

                for (int i = 0; i < WarningDataRow.Length; i++)
                {
                    WarningDataTable.ImportRow(WarningDataRow[i]);
                }

                for (int i = 0; i < ErrorRow.Length; i++)
                {
                    ErrorDataTable.ImportRow(ErrorRow[i]);
                }

                for (int i = 0; i < CriticalErrorRow.Length; i++)
                {
                    CriticalErrorDataTable.ImportRow(CriticalErrorRow[i]);
                }

                MergeSortDataTable = InformationDataTable.Copy();
                MergeSortDataTable.Merge(WarningDataTable, true);
                MergeSortDataTable.Merge(ErrorDataTable, true);
                MergeSortDataTable.Merge(CriticalErrorDataTable, true);

                Console.WriteLine("total " + MergeSortDataTable.Rows.Count.ToString() + " event");
            }
            catch(Exception e)
            {
                Console.WriteLine("MergeSortEventDataTable" + e.Message);
            }
        }

        public void ReadExcelFile(string Filepath, string SheetName, ref System.Data.DataTable EventDataTable)
        {
            try
            {
                string CommandString =
                    "Data Source=" + Filepath + ";" +
                    "Provider=" + ProviderName +
                    "Extended Properties=" + ExtendedString +
                    "HDR=" + Hdr +
                    "IMEX=" + IMEX;

                OleDbConnection DBConnection = new OleDbConnection(CommandString);
                DBConnection.Open();
                string QueryString = "select * from[" + SheetName + "$]";
                OleDbDataAdapter dr = new OleDbDataAdapter(QueryString, DBConnection);
                dr.Fill(EventDataTable);
                DBConnection.Close();

                Console.WriteLine("total " + EventDataTable.Rows.Count.ToString() + " event");
            }
            catch(Exception e)
            {
                Console.WriteLine("ReadExcelFile" + e.Message);
            }
        }

        /// <summary>
        /// 1. remove event that have no event ID
        /// 2. assign specific colume on new event datatable, maintenance can be easy
        /// </summary>
        public void ExcelDataPreProcessing(ref System.Data.DataTable InEventDataTable, ref System.Data.DataTable OutEventDataTable)
        {
            try
            {
                OutEventDataTable = InEventDataTable.Copy();

                //Assign Column name and handle colume that not valid
                for (int i = 0; i < OutEventDataTable.Columns.Count; i++)
                {
                    //if colume name is blank, fill index to be colume name
                    if (OutEventDataTable.Rows[0][i].ToString() == "")
                    {
                        OutEventDataTable.Rows[0][i] = i.ToString();
                        continue;
                    }
                    //if there exist the same colume name, fill index to be colume name
                    for (int j = 0; j < i;j++ )
                    {
                        if (OutEventDataTable.Rows[0][i].ToString() == OutEventDataTable.Rows[0][j].ToString())
                        {
                            OutEventDataTable.Rows[0][i] = i.ToString();
                        }
                    }
                }

                for (int i = 0; i < OutEventDataTable.Columns.Count; i++)
                {
                    OutEventDataTable.Columns[i].ColumnName = OutEventDataTable.Rows[0][i].ToString();
                }


                //start from i=1 since i=0 is colume name
                for (int i = 1; i < OutEventDataTable.Rows.Count; i++)
                {
                    string test = OutEventDataTable.Rows[i][1].ToString();

                    //if eventcode is blank, remove it
                    if (OutEventDataTable.Rows[i][1].ToString() == "")
                    {;
                        OutEventDataTable.Rows[i].Delete();
                        continue;
                    }
                }

                //delete first row as its event colume name
                OutEventDataTable.Rows[0].Delete();
                OutEventDataTable.AcceptChanges();

                Console.WriteLine("total " + OutEventDataTable.Rows.Count.ToString() + " event after preprocessing");

            }
            catch (Exception e)
            {
                Console.WriteLine("ExcelDataPreProcessing" + e.Message);
            }
        }

        public void WriteWordFile(string OutputFilePath, ref System.Data.DataTable EventDataTable)
        {
            try
            {
                /* 依照severity來做分類 
                    * Information: 8列 (差在action)
                    * Warning: 9列
                    * Error: 9列
                    * Critical Error: 9列
                    */
                int TableCount = EventDataTable.Rows.Count;
                //int TableCount = 130;

                Microsoft.Office.Interop.Word._Application word_app = new Microsoft.Office.Interop.Word.Application();
                Microsoft.Office.Interop.Word._Document word_document;
                object path;//設定一些object宣告
                object oMissing = System.Reflection.Missing.Value;
                object oSaveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdSaveChanges;
                object oformat = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatDocument;//wdFormatDocument97為Word 97-2003 文件 (*.doc)
                object start = 0, end = 0;
				int table_num = 7;

                path = OutputFilePath;
                word_app.Visible = false;//不顯示word程式
                word_app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;//不顯示警告或彈跳視窗。如果出現彈跳視窗，將選擇預設值繼續執行。
                word_document = word_app.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);//新增檔案
                Microsoft.Office.Interop.Word.Range rng = word_document.Range(ref start, ref end);

                Microsoft.Office.Interop.Word.Table table = word_document.Tables.Add(rng, table_num * TableCount, 2, ref oMissing, ref oMissing);//設定表格
                table.Select();//選取指令
                word_app.Selection.Font.Name = "Arial";//設定選取的資料字型
                word_app.Selection.Font.Size = 10;//設定文字大小
                //word_app.Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter; //將其設為靠中間
                word_app.Selection.Cells.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalTop; //將其設為靠上面
				try {//初始化表格
					
					for (int i = 1; i <= table_num * TableCount; i++)
					{
						for (int j = 1; j <= 2; j++)
						{
							if (i % table_num == 1)
							{
								table.Cell(i, j).Range.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorGray20;  //color of first Row
							}
							//table.Cell(i, j).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
							table.Cell(i, j).Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
						}
						table.Rows[i].Height = 10f;
						if (i % table_num == 0)
						{
							Console.Write("\r" + (i/table_num).ToString() + " / " + TableCount);
						}
					}
					table.Columns[1].Width = 90f;
					table.Columns[2].Width = 340f;
					Console.WriteLine();
				}
				catch (Exception e)
				{
					Console.WriteLine("inital table fail" + e.Message);
				}

                //填入文字
                for (int i = 1; i <= table_num * TableCount; i++)//將表格的資料寫入word檔案裡,第一列的值為1,第一行的值為1
                {
                    if (i % table_num == 0)
                    {
                        Console.Write("\r" + (i / table_num).ToString() + " / " + TableCount);
                    }
                    for (int j = 1; j <= 2; j++)
                    {
                        if (j == 1)
                        {
                            switch (i % table_num)
                            {
                                case 1:
                                    table.Cell(i, j).Range.Text = "Event ID";
                                    continue;
                                case 2:
                                    table.Cell(i, j).Range.Text = "Severity";
                                    continue;
                                case 3:
                                    table.Cell(i, j).Range.Text = "Category-Module";
                                    continue;
                                case 4:
                                    table.Cell(i, j).Range.Text = "Message";
                                    continue;
                                case 5:
                                    table.Cell(i, j).Range.Text = "1st Line Message";
                                    continue;
                                case 6:
                                    table.Cell(i, j).Range.Text = "Cause";
                                    continue;
                                case 0:
                                    table.Cell(i, j).Range.Text = "Action";
                                    continue;
                            }
                        }
                        else
                        {
                            switch (i % table_num)
                            {
                                case 1:
                                    string eventID = "#"+EventDataTable.Rows[i / table_num][0].ToString();
                                    table.Cell(i, j).Range.Hyperlinks.Add(table.Cell(i, j).Range, Address: eventID, TextToDisplay: EventDataTable.Rows[i / table_num][0].ToString());
                                    continue;
                                case 2:
                                    table.Cell(i, j).Range.Text = EventDataTable.Rows[i / table_num][4].ToString();
                                    continue;
                                case 3:
                                    table.Cell(i, j).Range.Text = EventDataTable.Rows[i / table_num][2].ToString();
                                    continue;
                                case 4:
                                    table.Cell(i, j).Range.Text = EventDataTable.Rows[i / table_num][7].ToString();
                                    continue;
                                case 5:
                                    table.Cell(i, j).Range.Text = EventDataTable.Rows[i / table_num][8].ToString();
                                    continue;
                                case 6:
                                    table.Cell(i, j).Range.Text = EventDataTable.Rows[i / table_num][5].ToString();
                                    continue;
                                case 0:
                                    table.Cell(i, j).Range.Text = EventDataTable.Rows[(i / table_num - 1)][6].ToString();
                                    continue;
                            }
                        }
                    }
                }
                Console.WriteLine();

                word_document.SaveAs(ref path, ref oformat, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing
                , ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);//存檔
                word_document.Close(ref oMissing, ref oMissing, ref oMissing);//關閉
                System.Runtime.InteropServices.Marshal.ReleaseComObject(word_document);//釋放
                word_document = null;
                word_app.Quit(ref oMissing, ref oMissing, ref oMissing);//結束
                System.Runtime.InteropServices.Marshal.ReleaseComObject(word_app);//釋放
                word_app = null;
            }
            catch (Exception e)
            {
                Console.WriteLine("WriteWordFile" + e.Message);
            }
        }

    }
}
