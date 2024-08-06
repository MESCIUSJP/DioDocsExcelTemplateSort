// See https://aka.ms/new-console-template for more information
using GrapeCity.Documents.Excel;
using System.Data;

Console.WriteLine("DioDocs for Excelのソート機能（テンプレート構文）");

// 新規ワークブックの作成
Workbook workbook = new();

// 帳票テンプレートを読み込む
//workbook.Open("test-0.xlsx"); // ソートなし
//workbook.Open("test-1.xlsx"); // C列でソート
workbook.Open("test-2.xlsx"); // C、F列でソート

#region データの初期化
DataTable datasource = new();
datasource.Columns.Add(new DataColumn("Id", typeof(int)));
datasource.Columns.Add(new DataColumn("Name", typeof(string)));
datasource.Columns.Add(new DataColumn("JapaneseLanguage", typeof(double)));
datasource.Columns.Add(new DataColumn("Mathematics", typeof(double)));
datasource.Columns.Add(new DataColumn("English", typeof(double)));
datasource.Columns.Add(new DataColumn("Total", typeof(double)));

datasource.Rows.Add(1, "小林 さゆり", 80, 90, 70, 240);
datasource.Rows.Add(2, "遠藤 花子", 70, 90, 90, 250);
datasource.Rows.Add(3, "吉田 洋介", 90, 80, 90, 260);
datasource.Rows.Add(4, "藤田 直人", 70, 70, 80, 220);
datasource.Rows.Add(5, "伊藤 七夏", 70, 90, 70, 230);
datasource.Rows.Add(6, "後藤 拓真", 70, 70, 80, 220);
datasource.Rows.Add(7, "藤井 舞", 70, 80, 90, 240);
datasource.Rows.Add(8, "伊藤 明美", 90, 90, 90, 270);
datasource.Rows.Add(9, "田中 さゆり", 80, 90, 70, 240);
datasource.Rows.Add(10, "吉田 直樹", 70, 70, 90, 230);
datasource.Rows.Add(11, "村上 里佳", 70, 80, 70, 220);
datasource.Rows.Add(12, "石川 翔太", 70, 90, 70, 230);
datasource.Rows.Add(13, "村上 治", 90, 90, 70, 250);
datasource.Rows.Add(14, "藤原 さゆり", 90, 90, 90, 270);
datasource.Rows.Add(15, "小林 聡太郎", 90, 90, 90, 270);
datasource.Rows.Add(16, "井上 直人", 70, 80, 70, 220);
datasource.Rows.Add(17, "清水 千代", 90, 80, 70, 240);
datasource.Rows.Add(18, "中島 浩", 70, 70, 70, 210);
datasource.Rows.Add(19, "佐藤 香織", 90, 80, 80, 250);
datasource.Rows.Add(20, "高橋 充", 90, 90, 70, 250);
#endregion

// データソースを追加
workbook.AddDataSource("ds", datasource);

// データを連結して帳票を作成
workbook.ProcessTemplate();

// Excelファイルに保存
workbook.Save("result.xlsx");