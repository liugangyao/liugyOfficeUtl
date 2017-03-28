using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace liugyOfficeUtl
{
    public class LiugyExcel
    {
        // *******************************************
        // copy from http://c-sharp-guide.com/?p=185
        // *******************************************

        /// <summary>
        /// Excelブックを新規に作成してデータテーブルの内容を出力します
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <param name="path">書き込み先のディレクトリパス</param>
        /// <param name="BookName">ファイル名称</param>
        /// <param name="SheetName">シート名称</param>
        /// 

        public static void ExcelWriter(System.Data.DataTable dt, string path, string BookName, string SheetName)
        {
            // Excelオブジェクトを生成
            Excel.Application ExcelApp = new Excel.Application();
            try
            {
                // 上書きの確認ダイアログを非表示
                ExcelApp.DisplayAlerts = false;
                // ウィンドウは非表示
                ExcelApp.Visible = false;
                // ブックを作成
                //Workbook wb = ExcelApp.Workbooks.Add();
                // 1枚目のシートを選択
                //Worksheet ws = wb.Sheets[SheetName];

                // ファイル属性を取得
                FileAttributes fas = File.GetAttributes(path + "\\" + BookName);
                // 読み取り専用かどうか確認
                if ((fas & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                {
                    Console.WriteLine("読み取り専用です。");

                    // ファイル属性から読み取り専用を削除
                    fas = fas & ~FileAttributes.ReadOnly;
                }

                // / ファイルオープンからワークブックの作成
                Excel.Workbook wb = ExcelApp.Workbooks.Open(path + "\\" + BookName);

                //シート名を指定して、インデックスを返す
                int sheetId = getSheetIndex(SheetName, wb);

                Excel.Worksheet ws = wb.Sheets[sheetId];

                ws.Select(Type.Missing);

                // データテーブルの列ごとでExcelに出力
                for (int col = 0; col < dt.Columns.Count; col++)
                {
                    // +1は項目名の行
                    object[,] obj = new object[dt.Rows.Count + 1, 1];

                    // 項目名出力
                    obj[0, 0] = dt.Columns[col].ColumnName;

                    for (int row = 0; row < dt.Rows.Count; row++)
                    {
                        // データテーブルをobject配列に格納
                        obj[row + 1, 0] = dt.Rows[row][col].ToString();
                    }

                    Excel.Range rgn = ws.Range[ws.Cells[1, col + 1], ws.Cells[dt.Rows.Count + 1, col + 1]]; //最初のセル(行,列)～最後のセル(行,列)
                   // rgn.Font.Size = 10;
                   // rgn.Font.Name = "メイリオ";

                    DataColumn dtcol = dt.Columns[col];
                    if (dtcol.DataType.ToString() == "System.String")
                    {
                        rgn.NumberFormatLocal = "@";  // 表示形式を文字列にする
                        rgn.Value2 = obj;
                    }
                    else //System.Int32
                    {
                        rgn.Value2 = obj;
                    }
                }

                // Bookを保存
                wb.SaveAs(path +"\\"+ BookName);
                // Bookを閉じる
                wb.Close(false);
                // Excelアプリケーションの終了
                ExcelApp.Quit();
                // COMの解放
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
            }
            catch (Exception ex)
            {
             //   Console.Write(BookName + "を作成できませんでした。");
             //   Console.Write(ex.Message);
                // Excelアプリケーションの終了
                ExcelApp.Quit();
                throw ex;
            }

        }


        /// <summary>
        /// Excel データをDataTableに格納します
        /// </summary>
        /// <param name="path">Excel格納ディレクトリ</param>
        /// <param name="BookName">ファイル名称</param>
        /// <param name="SheetName">シート名称</param>
        /// <return>読み込んだExcelデータ</return>
        public static System.Data.DataTable ExcelReader(string path, string BookName, string SheetName)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            // Excelオブジェクトを生成
            Excel.Application ExcelApp = new Excel.Application();
            try
            {
                // ウィンドウは非表示
                ExcelApp.Visible = false;
                // エクセルオープン
                Excel.Workbook wb = ExcelApp.Workbooks.Open(path +"\\"+ BookName);

                //シート名を指定して、インデックスを返す
                int sheetId = getSheetIndex(SheetName, wb);

                Excel.Worksheet sheet = wb.Sheets[sheetId];

                //1シート目の選択
                //Excel.Worksheet sheet = WorkBook.Sheets[1];
                sheet.Select();

                // 最大行数
                int Maxrow = sheet.get_Range("A1").End[Excel.XlDirection.xlDown].Row;
                // 最大列数
                int Maxcol = sheet.UsedRange.Columns.Count;

                for (int i = 0; i < Maxcol; i++)
                {
                    //カラム名にダミーを設定します。
                    dt.Columns.Add("ダミー" + i);
                }

                for (int col = 0; col < Maxcol; col++)
                {
                    Excel.Range rg = sheet.Cells[1, col + 1];
                    dt.Columns[col].ColumnName = rg.Value2;
                }

                Excel.Range Excel_data = sheet.get_Range("A2", Type.Missing).get_Resize(Maxrow, Maxcol);
                for (int row = 2; row <= Maxrow; row++)
                {
                    DataRow dr = dt.NewRow();
                    for (int col = 1; col <= Maxcol; col++)
                    {
                        Excel_data = sheet.Cells[row, col];
                        dr[dt.Columns[col - 1].ColumnName] = Excel_data.Value2;
                    }
                    dt.Rows.Add(dr);
                }

                //workbookを閉じる
                wb.Close();
                //エクセルを閉じる
                ExcelApp.Quit();
                // COMの解放
                System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp);
            }
            catch (Exception ex)
            {
                Console.Write("Excelファイルを読み込めませんでした。");
                //エクセルを閉じる
                ExcelApp.Quit();
            }
            return dt;
        }

        //シート名から：アクセスインデックスを探し出す
        private static int getSheetIndex(string sheetName, Excel.Workbook m_workBook)
        {
            int i = 0;
            foreach (Excel.Worksheet sh in m_workBook.Sheets)
            {
                if (sheetName == sh.Name)
                {
                    return i + 1;
                }
                i += 1;
            }
            return -1;
        }

    }

}
