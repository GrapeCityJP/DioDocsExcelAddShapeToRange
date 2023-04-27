// See https://aka.ms/new-console-template for more information
using GrapeCity.Documents.Common;
using GrapeCity.Documents.Excel;
using GrapeCity.Documents.Excel.Drawing;
using System.Drawing;

Console.WriteLine("指定したセル範囲に図形や画像、チャートを追加する");

var workbook = new Workbook();

#region 図形を追加
var worksheet1 = workbook.Worksheets[0];
worksheet1.Name = "図形";

// 図形を追加
//var rect = CellInfo.GetAccurateRangeBoundary(worksheet1.Range["C4:G5"]);
//IShape shape = worksheet1.Shapes.AddShapeInPixel(AutoShapeType.Rectangle, rect.Left, rect.Top, rect.Width, rect.Height);
//shape.Line.Visible = false;
//shape.Fill.Color.ObjectThemeColor = ThemeColor.Accent1;
//shape.TextFrame.TextRange.Text = "DioDocs（ディオドック）";
//shape.TextFrame.TextRange.Font.Size = 20;
//shape.TextFrame.TextRange.Font.Name = "Calibri";
//shape.TextFrame.TextRange.Font.Color.RGB = Color.White;
//shape.TextFrame.HorizontalAnchor = HorizontalAnchor.Center;
//shape.TextFrame.VerticalAnchor = VerticalAnchor.AnchorMiddle;

// 図形を追加（V6J）
IShape shape = worksheet1.Shapes.AddShape(AutoShapeType.Rectangle, worksheet1.Range["C4:G5"]);
shape.Line.Visible = false;
shape.Fill.Color.ObjectThemeColor = ThemeColor.Accent1;
shape.TextFrame.TextRange.Text = "DioDocs（ディオドック）";
shape.TextFrame.TextRange.Font.Size = 20;
shape.TextFrame.TextRange.Font.Name = "Calibri";
shape.TextFrame.TextRange.Font.Color.RGB = Color.White;
shape.TextFrame.HorizontalAnchor = HorizontalAnchor.Center;
shape.TextFrame.VerticalAnchor = VerticalAnchor.AnchorMiddle;

#endregion

#region 画像を追加

var worksheet2 = workbook.Worksheets.Add();
worksheet2.Name = "画像";

// 画像を追加
//var rect = CellInfo.GetAccurateRangeBoundary(worksheet2.Range["C4:F6"]);
//IShape picture = worksheet2.Shapes.AddPictureInPixel("DioDocs.png", rect.Left, rect.Top, rect.Width, rect.Height);

// 画像を追加（V6J）
IShape picture = worksheet2.Shapes.AddPicture("DioDocs.png", worksheet2.Range["C4:F6"]);

#endregion

#region チャートを追加

var worksheet3 = workbook.Worksheets.Add();
worksheet3.Name = "チャート";

// データ
var data = new object[,]
{
    {"国・地域", "第1四半期", "第2四半期", "第3四半期", "第4四半期" },
    {"オーストラリア", 16439, 18106, 15193, 14879},
    {"中国", 42659, 14392, 42284, 38270},
    {"日本", 44000, 15039, 27961, 34382},
    {"アメリカ", 23174, 42797, 23637, 26200}
};

worksheet3.Range["A1:E5"].Value = data;
worksheet3.Range["A1:E5"].AutoFit();
worksheet3.Range["B2:E5"].NumberFormat = @"¥#,##0";

// チャートを追加
//var rect = CellInfo.GetAccurateRangeBoundary(worksheet3.Range["B7:L26"]);
//IShape chartCol = worksheet3.Shapes.AddChartInPixel(ChartType.ColumnClustered, rect.Left, rect.Top, rect.Width, rect.Height);
//chartCol.Chart.SeriesCollection.Add(worksheet3.Range["A1:E5"]);
//chartCol.Chart.ChartTitle.Text = "四半期売上";

// チャートを追加（V6J）
IShape chartCol = worksheet3.Shapes.AddChart(ChartType.ColumnClustered, worksheet3.Range["B7:L26"]);
chartCol.Chart.SeriesCollection.Add(worksheet3.Range["A1:E5"]);
chartCol.Chart.ChartTitle.Text = "四半期売上";

#endregion

// Excelファイルに保存
workbook.Save("AddShapeToRange.xlsx");