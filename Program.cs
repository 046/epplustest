using System;
using System.Collections.Generic;
using System.IO;

using OfficeOpenXml;

namespace xls
{
    class Program
    {
        // inputディレクトリのファイルの増幅倍率を入れる
        const int TIMES = 30;

        static void Main(string[] args)
        {
            // ライセンス設定
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            DateTime dtStart = DateTime.Now;

            // inputディレクトリのファイル一覧を取得
            var inputs = Directory.GetFiles(@"data\input");
            var files = new List<string>();
            for (int i = 0; i < TIMES; i++) {
                foreach (string file in inputs) {
                    files.Add(file);
                }
            }

            // 統合結果ファイル名
            string resultFile = @"data\output\result.xlsx";

            // シート数カウント用
            int sheetCnt = 0;

            // 空のExcelファイルを作る
            ExcelPackage masterPackage = new ExcelPackage();

            // 統合メイン処理
            foreach (var file in files)
            {
                ExcelPackage pckg = new ExcelPackage(new FileInfo(file));

                foreach (var sheet in pckg.Workbook.Worksheets)
                {
                    sheetCnt++;

                    string workSheetName = sheet.Name;

                    foreach (var masterSheet in masterPackage.Workbook.Worksheets)
                    {
                        if (sheet.Name == masterSheet.Name)
                        {
                            // 同じシート名があればマイクロ秒の時間をつける
                            workSheetName = string.Format("{0}_{1}", workSheetName, DateTime.Now.ToString("yyyyMMddhhmmssffffff"));
                        }

                    }

                    // シート追加
                    masterPackage.Workbook.Worksheets.Add(workSheetName, sheet);
                    Console.WriteLine($"シート{sheetCnt}: {workSheetName}");

                }
            }

            masterPackage.SaveAs(new FileInfo(resultFile));

            // 処理時間表示
            Console.WriteLine($"処理時間: {DateTime.Now - dtStart}");
        }
    }
}
