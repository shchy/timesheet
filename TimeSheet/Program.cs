using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;
using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

namespace TimeSheet
{
    using Excel = NetOffice.ExcelApi;

    class Program
    {
        static void Main(string[] args)
        {
            // todo 入力
            var planFilePath = @"C:\Users\shch\Downloads\BP勤怠チェック\CM2B1G体制表＆登録番号一覧（2020.2月）_r2.xlsx";
            var orderFilePath = @"C:\Users\shch\Downloads\BP勤怠チェック\2020年度_協力会社発注管理表(CM2B1G).xlsm";
            var timeSheetFolder = @"C:\Users\shch\Downloads\サンプル";

            using (var app = new Excel.Application { Visible = false, DisplayAlerts = false, })
            {
                // 最初に1枚にまとめる
                var mergedBook = MergeBooks(app, planFilePath, orderFilePath);

                // 発注管理表を読み込む
                var orderSheet = mergedBook.Worksheets.OfType<Worksheet>().First(x => x.Name.Contains("協力会社発注管理表"));
                var orderFile = new OrderFile();
                var orders = orderFile.Read(orderSheet).ToArray();

                // マージシートを作成
                var mergeSheet = mergedBook.Worksheets.Add() as Worksheet;
                mergeSheet.Move(mergedBook.Worksheets.First(), Type.Missing);

                // 勤怠シートを読み込む
                var works = new List<TimeSheet>();
                var timeSheetFile = new TimeSheetFile();
                var files = Directory.GetFiles(timeSheetFolder)
                    .Where(x => Path.GetFileName(x).EndsWith(".xlsx"))
                    .Where(x => !Path.GetFileName(x).StartsWith("~"));
                foreach (var path in files)
                {
                    using (var timeBook = app.Workbooks.Open(path))
                    {
                        // 勤怠シートをコピー
                        var timeSheet = timeBook.Worksheets.OfType<Worksheet>().FirstOrDefault(x => x.Name.Contains("C322"));
                        if (timeSheet == null)
                        {
                            throw new Exception($"{path}にC322シートが見つかりませんでした");
                        }

                        // モデル化
                        var work = timeSheetFile.Read(timeSheet);
                        works.Add(work);

                        // コピー先の指定がないと新しいBookが作られる
                        timeSheet.Name = $"勤怠({work.Worker.Name})";
                        timeSheet.Copy(Type.Missing, mergedBook.Worksheets.Last());
                        timeBook.Close();
                    }
                }


                MakeListSheet(mergedBook, mergeSheet, orders, works);
                // todo 保存先
                mergedBook.SaveAs(@"C:\Users\shch\Downloads\BP勤怠チェック\asdf.xlsx");

                foreach (var book in app.Workbooks)
                {
                    book.Close();
                }
            }
        }




        private static void MakeListSheet(
            Workbook book,
            Worksheet sheet,
            IEnumerable<Order> orders,
            IEnumerable<TimeSheet> works)
        {
            sheet.Activate();
            // todo 
            sheet.Name = "一覧";

            var headers = new[] { 
                "協力会社名", 
                "要員名", 
                "人月", 
                "基準単価", 
                "下限時間", 
                "上限時間", 
                "控除単価", 
                "超過単価", 
                "PRJ番号", 
                "工数", 
                "精算金額",
                "経費(税込)",
                "経費",
                "検収金額", 
                "チェック結果" 
            };
            
            var header = sheet.Range("A1:O1");
            header.Value = headers;

            var y = 2;
            foreach (var order in orders)
            {
                var linkWorks = (
                    from work in works
                    where WSClear(work.Worker.Name) == WSClear(order.Name)
                    where WSClear(work.Worker.OrderID) == WSClear(order.OrderID)
                    where work.WorkRecords.Any()
                    let workMonth = work.WorkRecords.First().Date
                    where workMonth.Year == order.Date.Year
                    where workMonth.Month == order.Date.Month
                    select work
                    ).ToArray();
                var underTime = order.WorkPerMonth * order.UnderTime.TotalHours;
                var upperTime = order.WorkPerMonth * order.UnderTime.TotalHours;

                foreach (var work in linkWorks)
                {
                    foreach (var assign in work.WorkAssigns)
                    {
                        sheet.Rows[y].Cells[1].Value = order.Company;
                        sheet.Rows[y].Cells[2].Value = order.Name;
                        sheet.Rows[y].Cells[3].Value = order.WorkPerMonth;
                        sheet.Rows[y].Cells[4].Value = order.Price;
                        sheet.Rows[y].Cells[5].Value = order.UnderTime.TotalHours;
                        sheet.Rows[y].Cells[6].Value = order.UpperTime.TotalHours;
                        sheet.Rows[y].Cells[7].Value = order.UnderPriceByHour;
                        sheet.Rows[y].Cells[8].Value = order.UpperPriceByHour;
                        sheet.Rows[y].Cells[9].Value = assign.ID;
                        sheet.Rows[y].Cells[10].Value = assign.Value.TotalHours;
                        var payOff = 0.0;
                        if (assign.Value.TotalHours < underTime)
                        {
                            payOff = (assign.Value.TotalHours - underTime) * order.UnderPriceByHour;
                        }
                        else if (upperTime < assign.Value.TotalHours)
                        {
                            payOff = (assign.Value.TotalHours - upperTime) * order.UpperPriceByHour;
                        }
                        sheet.Rows[y].Cells[11].Value = payOff;
                        sheet.Rows[y].Cells[12].Value = assign.Expense;
                        var withoutTax = Math.Floor(assign.Expense / 1.1);
                        sheet.Rows[y].Cells[13].Value = withoutTax;
                        sheet.Rows[y].Cells[14].Value = order.Price + payOff + withoutTax;
                        y++;
                    }
                }
            }
        }

        static string WSClear(string x1)
        {
            return x1.Replace(" ", "").Replace("　","").Trim();
        }

        private static Workbook MergeBooks(Application app,
                                           string planFilePath,
                                           string orderFilePath)
        {
            using (var planBook = app.Workbooks.Open(planFilePath)) {
                // 体制表から直接シートをコピー
                var planSheet = planBook.Worksheets.OfType<Worksheet>().FirstOrDefault(x => x.Name.Contains("直接"));
                if (planSheet == null)
                {
                    throw new Exception($"{planFilePath}に直接シートが見つかりませんでした");
                }
                // コピー先の指定がないと新しいBookが作られる
                planSheet.Copy(Type.Missing, Type.Missing);
                planBook.Close();
            }

            // 残ったBookがマージ先Book
            var mergedBook = app.Workbooks.First();

            using (var orderBook = app.Workbooks.Open(orderFilePath)) {

                // 発注管理から協力会社発注管理表シートをコピー
                var orderSheet = orderBook.Worksheets.OfType<Worksheet>().FirstOrDefault(x => x.Name.Contains("協力会社発注管理表"));
                if (orderSheet == null)
                {
                    throw new Exception($"{orderFilePath}に協力会社発注管理表シートが見つかりませんでした");
                }
                // コピー先の指定がないと新しいBookが作られる
                orderSheet.Copy(Type.Missing, mergedBook.Worksheets.Last());
                orderBook.Close();
            }

            return mergedBook;
        }
    }
}
