using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using OpenXmlDemo.Api.Help;
using System.Security.Policy;

namespace OpenXmlDemo.Api.Controllers
{
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class ExportController : ControllerBase
    {

        private readonly string _rootPath;
        private readonly float _resolution = 96;
        private readonly string _picAddress = "/static/images/1.jpg";
        private readonly string _url = "http://localhost:5000";

        public ExportController(IWebHostEnvironment environment)
        {
            _rootPath = environment?.WebRootPath;
        }

        [HttpGet]
        public async Task<IActionResult> ExportWord()
        {
            string savePath = Path.Combine(_rootPath, "Word");
            if (!Directory.Exists(savePath))
                Directory.CreateDirectory(savePath);
            string fileName = $"{Guid.NewGuid()}.docx";
            using var doc = WordprocessingDocument.Create(Path.Combine(savePath, fileName), WordprocessingDocumentType.Document);
            // 新增主文档部分
            var main = doc.AddMainDocumentPart();
            main.Document = new Document();
            #region 第一页
            // 新增主体
            var docBody1 = main.Document.AppendChild(new Body());
            // 设置主体页面大小及方向
            docBody1.AppendChild(new SectionProperties(new PageSize { Width = 11906U, Height = 16838U, Orient = PageOrientationValues.Portrait }));

            docBody1.Wrap(2, "72");
            docBody1.SetParagraph("藏品档案", fontSize: "72", isBold: true, wrapNum: 2);
            docBody1.SetParagraph($"文物名称：", fontSize: "36", justification: 0, wrapNum: 2);
            docBody1.SetParagraph($"藏品级别：", fontSize: "36", justification: 0, wrapNum: 2);
            docBody1.SetParagraph($"总登记号：", fontSize: "36", justification: 0, wrapNum: 2);
            docBody1.SetParagraph($"分 类 号：", fontSize: "36", justification: 0, wrapNum: 2);
            docBody1.SetParagraph($"档案编号：", fontSize: "36", justification: 0, wrapNum: 2);
            docBody1.SetParagraph($"收藏单位：", fontSize: "36", justification: 0, wrapNum: 2);
            docBody1.SetParagraph($"制档日期：{DateTime.Now:yyyy年MM月dd日}", fontSize: "36", justification: 0, wrapNum: 2);
            docBody1.SetParagraph($"制 档 人：", fontSize: "36", justification: 0, wrapNum: 2);
            #endregion

            int initRow = 34, initCol = 8;
            string topMargion = "100", bottomMargion = "100";

            #region 第二页到第四页

            var docBody2 = main.Document.AppendChild(new Body());
            docBody2.CreatePortraitSectionBreakParagraph();
            docBody2.SetParagraph("文物藏品登记表", fontSize: "36");
            Table baseTable = docBody2.AppendChild(new Table());
            baseTable.InitTable(initRow, initCol, 1, topMargion: topMargion, bottomMargion: bottomMargion);
            for (int row = 0; row < initRow; row++)
            {
                var tabRow = baseTable.Descendants<TableRow>().ElementAt(row);
                // 自定义行高
                List<int> heightArr = new() { 10, 14, 15, 17, 20, 21, 24, 25, 26, 27, 28, 29, 30, 32, 33 };
                if (heightArr.Any(d => d == row))
                {
                    var rowProp = tabRow?.Elements<TableRowProperties>()?.First();
                    var rowHeight = rowProp.Elements<TableRowHeight>()?.FirstOrDefault();
                    switch (row)
                    {
                        case 10:
                        case 28:
                        case 32:
                        case 33:
                            rowHeight.Val = 3600;
                            break;
                        case 17:
                            rowHeight.Val = 1900;
                            break;
                        case 14:
                        case 15:
                        case 20:
                        case 21:
                            rowHeight.Val = 2200;
                            break;
                        case 24:
                        case 25:
                        case 26:
                        case 27:
                            rowHeight.Val = 2800;
                            break;
                    }
                }
                var tabcells = tabRow.Descendants<TableCell>();
                for (int col = 0; col < initCol; col++)
                {
                    var tabCell = tabcells.ElementAt(col);
                    var tabCellProps = tabCell.AppendChild(new TableCellProperties());
                    var tabCellPara = tabCell.Elements<Paragraph>().First();
                    var tableCellParaPorp = tabCellPara.AppendChild(new ParagraphProperties());
                    var tabCellRun = tabCellPara.AppendChild(new Run());
                    // 设置单元格字体、大小、颜色
                    var runProperties = tabCellRun.AppendChild(new RunProperties());
                    string text = string.Empty;
                    if (col == 0)
                    {
                        switch (row)
                        {
                            case 0:
                                text = "名称";
                                break;
                            case 1:
                                text = "曾用名";
                                break;
                            case 2:
                                text = "总登记号";
                                break;
                            case 3:
                                text = "入馆登记号";
                                break;
                            case 4:
                                text = "分类账号";
                                break;
                            case 5:
                                text = "类别";
                                break;
                            case 6:
                                text = "年代类型";
                                break;
                            case 7:
                                text = "年代研究信息";
                                break;
                            case 8:
                                text = "地域类型";
                                break;
                            case 9:
                                text = "人文类型";
                                break;
                            case 10:
                                text = "人物传略";
                                break;
                            case 11:
                                text = "质地";
                                break;
                            case 12:
                                text = "尺寸";
                                break;
                            case 13:
                                text = "传统数量";
                                break;
                            case 14:
                                text = "形态特征";
                                break;
                            case 15:
                                text = "工艺技法";
                                break;
                            case 16:
                                text = "完残程度";
                                break;
                            case 17:
                                text = "完残状况";
                                break;
                            case 18:
                                text = "颜色";
                                break;
                            case 19:
                                text = "文字种类";
                                break;
                            case 20:
                                text = "题识情况";
                                break;
                            case 21:
                                text = "附属物情况";
                                break;
                            case 22:
                                text = "来源方式";
                                break;
                            case 23:
                                text = "来源单位或个人";
                                break;
                            case 24:
                                text = "搜集经过";
                                break;
                            case 25:
                                text = "流传经历";
                                break;
                            case 26:
                                text = "出土情况";
                                break;
                            case 27:
                                text = "鉴定情况";
                                break;
                            case 28:
                                text = "当前状况";
                                break;
                            case 29:
                                text = "保存条件";
                                break;
                            case 30:
                                text = "损坏原因";
                                break;
                            case 31:
                                text = "优先保护等级";
                                break;
                            case 32:
                                text = "历史保护记录";
                                break;
                            case 33:
                                text = "主要利用情况记录";
                                break;
                        }
                    }
                    else
                    {
                        // 设置title
                        switch (col)
                        {
                            case 4 when (row < 12 || row == 16 || row == 18 || row == 22 || row == 31):
                                if (row == 2)
                                    text = "入藏日期";
                                else if (row == 3)
                                    text = "入馆日期";
                                else if (row == 4)
                                    text = "入藏库房";
                                else if (row == 5)
                                    text = "级别";
                                else if (row == 6)
                                    text = "年代";
                                else if (row == 8)
                                    text = "地域";
                                else if (row == 9)
                                    text = "人文";
                                else if (row == 11)
                                    text = "功能类别";
                                else if (row == 16)
                                    text = "独特标记";
                                else if (row == 18)
                                    text = "光泽";
                                else if (row == 22)
                                    text = "来源号";
                                else if (row == 31)
                                    text = "拟采取的保护措施";
                                break;
                            default:
                                if (row == 13 && col == 2)
                                    text = "实际数量";
                                else if (row == 13 && col == 4)
                                    text = "容积";
                                else if (row == 13 && col == 6)
                                    text = "质量";
                                else if (row == 19 && col == 2)
                                    text = "字体";
                                else if (row == 19 && col == 4)
                                    text = "字迹颜色";
                                break;
                        }
                        // 设置value
                        switch (row)
                        {
                            case 0:
                            case 1:
                            case 7:
                            case 10:
                            case 12:
                            case 14:
                            case 15:
                            case 17:
                            case 20:
                            case 21:
                            case 23:
                            case 24:
                            case 25:
                            case 26:
                            case 27:
                            case 28:
                            case 29:
                            case 30:
                            case 32:
                            case 33:
                                text = "";
                                if (col == 1)
                                {
                                    tabCellProps.AppendChild(new HorizontalMerge() { Val = MergedCellValues.Restart });
                                }
                                else
                                {
                                    tabCellProps.AppendChild(new HorizontalMerge() { Val = MergedCellValues.Continue });
                                }
                                break;
                            case 2:
                            case 3:
                            case 4:
                            case 5:
                            case 6:
                            case 8:
                            case 9:
                            case 11:
                            case 16:
                            case 18:
                            case 22:
                            case 31:
                                if (col < 4)
                                {
                                    text = "";
                                }
                                else if (col > 4)
                                {
                                    text = "";
                                }
                                if (col == 1 || col == 5)
                                {
                                    tabCellProps.AppendChild(new HorizontalMerge() { Val = MergedCellValues.Restart });
                                }
                                else if ((col > 1 && col < 4) || (col > 4 && col < initCol))
                                {
                                    tabCellProps.AppendChild(new HorizontalMerge() { Val = MergedCellValues.Continue });
                                }
                                break;
                            case 13:
                                text = "";
                                break;
                            case 19:
                                if (col == 1)
                                    text = "";
                                else if (col == 3)
                                    text = "";
                                else if (col == 5)
                                {
                                    text = "";
                                    tabCellProps.AppendChild(new HorizontalMerge() { Val = MergedCellValues.Restart });
                                }
                                else if (col < initCol)
                                    tabCellProps.AppendChild(new HorizontalMerge() { Val = MergedCellValues.Continue });
                                break;
                            default:
                                break;
                        }
                    }
                    tabCellRun.AppendChild(new Text(text));
                    tableCellParaPorp.AppendChild(new Justification() { Val = JustificationValues.Left });
                    runProperties.AppendChild(new RunFonts { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体" });
                    runProperties.AppendChild(new FontSize { Val = "18" });
                    runProperties.AppendChild(new Color { Val = "#000000" });
                }
            }
            #endregion

            #region 第六页

            var docBody6 = main.Document.AppendChild(new Body());
            docBody6.CreatePortraitSectionBreakParagraph();
            initRow = 1; initCol = 2;
            Table baseTable6 = docBody6.AppendChild(new Table());
            baseTable6.InitTable(initRow, initCol, 1, topMargion: topMargion, bottomMargion: bottomMargion);
            for (int row = 0; row < initRow; row++)
            {
                var tabRow = docBody6.Descendants<TableRow>().ElementAt(row);
                var tabcells = tabRow.Descendants<TableCell>();
                for (int col = 0; col < initCol; col++)
                {
                    var rowProp = tabRow?.Elements<TableRowProperties>()?.First();
                    var rowHeight = rowProp.Elements<TableRowHeight>()?.FirstOrDefault();
                    rowHeight.Val = 13000U;
                    var tabCell = tabcells.ElementAt(col);
                    var tabCellProps = tabCell.AppendChild(new TableCellProperties());
                    var tabCellPara = tabCell.Elements<Paragraph>().First();
                    var tableCellParaPorp = tabCellPara.AppendChild(new ParagraphProperties());
                    var tabCellRun = tabCellPara.AppendChild(new Run());
                    // 设置单元格字体、大小、颜色
                    var runProperties = tabCellRun.AppendChild(new RunProperties());
                    var cellWidth = tabCellProps.Elements<TableCellWidth>().FirstOrDefault();
                    if (cellWidth == null)
                        cellWidth = tabCellProps.AppendChild(new TableCellWidth());
                    if (col == 0)
                    {
                        cellWidth.Width = "500";
                        tabCellRun.AppendChild(new Text("附页"));
                        tabCellProps.AppendChild(new TextDirection() { Val = TextDirectionValues.TopToBottomLeftToRightRotated });
                        tableCellParaPorp.AppendChild(new Justification() { Val = JustificationValues.Center });
                        tabCellProps.AppendChild(new VerticalMerge() { Val = MergedCellValues.Restart });
                    }
                    else
                    {
                        string imgPath = await FileHelper.Down(_url + "/static/images/2.jpg");
                        string picType = imgPath.Split('.').Last().ToLower();
                        picType = picType == "jpg" ? "jpeg" : picType;
                        ImagePart imagePart = null;
                        if (Enum.TryParse(picType, true, out ImagePartType imagePartType))
                        {
                            imagePart = main.AddImagePart(imagePartType);
                        }
                        var fs = System.IO.File.Open(imgPath, FileMode.Open);
                        imagePart?.FeedData(fs);
                        System.Drawing.Image image = System.Drawing.Image.FromStream(fs);
                        double img_width = image.Width;
                        double img_height = image.Height;
                        while (Math.Round(img_width / _resolution * 2.54, 1) > 5.5 || Math.Round(img_height / _resolution * 2.54, 1) > 5.5)
                        {
                            img_width /= 2;
                            img_height /= 2;
                        }
                        tabCellRun.AddImageToBodyTableCell(main.GetIdOfPart(imagePart), (long)(img_width / _resolution * 2.54 * 360000), (long)(img_height / _resolution * 2.54 * 360000));
                        image.Dispose();
                        await fs.DisposeAsync();
                        cellWidth.Width = "4500";
                    }
                    runProperties.AppendChild(new RunFonts { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体" });
                    runProperties.AppendChild(new FontSize { Val = "18" });
                    runProperties.AppendChild(new Color { Val = "#000000" });
                }
            }
            #endregion

            #region 文物藏品档案专用纸（图片）

            var docBody8 = main.Document.AppendChild(new Body());
            docBody8.CreatePortraitSectionBreakParagraph();
            docBody8.SetParagraph("文物藏品档案专用纸", isBold: true);
            docBody8.SetParagraph("照片册页", fontSize: "24", justification: 0);


            await InsertImages(main, docBody8, _picAddress, _url, _resolution);

            initRow = 4; initCol = 4;
            Table baseTable8 = docBody8.AppendChild(new Table());
            baseTable8.InitTable(initRow, initCol, 0, topMargion: topMargion, bottomMargion: bottomMargion);
            for (int row = 0; row < initRow; row++)
            {
                var tabRow = baseTable8.Descendants<TableRow>().ElementAt(row);
                var tabcells = tabRow.Descendants<TableCell>();
                for (int col = 0; col < initCol; col++)
                {
                    var tabCell = tabcells.ElementAt(col);
                    var tabCellProps = tabCell.AppendChild(new TableCellProperties());
                    var tabCellPara = tabCell.Elements<Paragraph>().First();
                    var tableCellParaPorp = tabCellPara.AppendChild(new ParagraphProperties());
                    var tabCellRun = tabCellPara.AppendChild(new Run());
                    // 设置单元格字体、大小、颜色
                    var runProperties = tabCellRun.AppendChild(new RunProperties());

                    string text = string.Empty;
                    if (col == 0)
                    {
                        if (row == 0)
                            text = "提名";
                        else if (row == 1)
                            text = "数码照片编号/底片号";
                        else if (row == 2)
                            text = "摄影者";
                        else
                            text = "说明";
                    }
                    else
                    {
                        if (row == 0 && col == 1)
                        {
                            text = "";
                            tabCellProps.AppendChild(new HorizontalMerge() { Val = MergedCellValues.Restart });
                        }
                        else if (row == 0 && col > 1)
                            tabCellProps.AppendChild(new HorizontalMerge() { Val = MergedCellValues.Continue });
                        else if (row == 1 && col == 1)
                            text = "";
                        else if (row == 1 && col == 2)
                            text = "参见号";
                        else if (row == 1 && col == 3)
                            text = "";
                        else if (row == 2 && col == 1)
                            text = "";
                        else if (row == 2 && col == 2)
                            text = "摄影日期";
                        else if (row == 2 && col == 3)
                            text = "";
                        else if (row == initRow - 1 && col == 1)
                        {
                            text = "";
                            tabCellProps.AppendChild(new HorizontalMerge() { Val = MergedCellValues.Restart });
                        }
                        else if (row == initRow - 1 && col > 1)
                            tabCellProps.AppendChild(new HorizontalMerge() { Val = MergedCellValues.Continue });

                    }
                    tabCellRun.AppendChild(new Text(text));
                    runProperties.AppendChild(new RunFonts { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体" });
                    runProperties.AppendChild(new FontSize { Val = "24" });
                    runProperties.AppendChild(new Color { Val = "#000000" });
                }
            }

            #endregion

            #region 备考表

            var docBody13 = main.Document.AppendChild(new Body());
            docBody13.CreatePortraitSectionBreakParagraph();
            docBody13.SetParagraph("备考表", isBold: true);

            initRow = 1;
            initCol = 1;
            Table baseTable13 = docBody13.AppendChild(new Table());
            baseTable13.InitTable(initRow, initCol, 1, rowHeight: 12000U, leftMargion: "200", topMargion: "200", bottomMargion: bottomMargion, cellAlignmentMethod: 0);

            for (int row = 0; row < initRow; row++)
            {
                var tabRow = docBody13.Descendants<TableRow>().ElementAt(row);
                var tabcells = tabRow.Descendants<TableCell>();
                for (int col = 0; col < initCol; col++)
                {
                    var tabCellPara = tabcells.ElementAt(col).Elements<Paragraph>().First();
                    var tabCellRun = tabCellPara.AppendChild(new Run());
                    var tableCellParaPorp = tabCellPara.AppendChild(new ParagraphProperties());

                    // 设置单元格字体、大小、颜色
                    var runProperties = tabCellRun.AppendChild(new RunProperties());

                    tabCellRun.CellWrap(1);

                    tabCellRun.AppendChild(new Text("说明："));

                    tabCellRun.CellWrap(8);

                    string beforeSpace = string.Empty;
                    string nextSpace = string.Empty;
                    for (int i = 0; i < 23; i++)
                    {
                        beforeSpace += "\r\n";
                    }

                    for (int i = 0; i < 28; i++)
                    {
                        nextSpace += "\r\n";
                    }

                    tabCellRun.AppendChild(new Text($@"{beforeSpace}立卷人："));
                    tabCellRun.CellWrap(1);
                    tabCellRun.AppendChild(new Text($@"{nextSpace}年  月  日"));
                    tabCellRun.CellWrap(1);
                    tabCellRun.AppendChild(new Text($@"{beforeSpace}检查人："));
                    tabCellRun.CellWrap(1);
                    tabCellRun.AppendChild(new Text($@"{nextSpace}年  月  日"));

                    tabCellRun.CellWrap(6);

                    tabCellRun.AppendChild(new Text("历次使用情况记录："));

                    runProperties.AppendChild(new RunFonts { Ascii = "宋体", HighAnsi = "宋体", EastAsia = "宋体" });
                    runProperties.AppendChild(new FontSize { Val = "24" });
                    runProperties.AppendChild(new Color { Val = "#000000" });
                }
            }

            #endregion

            return Ok(Path.Combine(savePath, fileName));
        }

        /// <summary>
        /// 插入图片
        /// </summary>
        /// <param name="main"></param>
        /// <param name="docBody"></param>
        /// <param name="picAddress"></param>
        /// <param name="url"></param>
        /// <param name="resolution"></param>
        /// <returns></returns>
        private static async Task InsertImages(MainDocumentPart main, Body docBody, string picAddress, string url, float resolution)
        {
            int initWidth = 16, initHeight = 20;
            if (!string.IsNullOrWhiteSpace(picAddress))
            {
                picAddress = await FileHelper.Down(url + picAddress);
                if (!string.IsNullOrWhiteSpace(picAddress))
                {
                    string picType = picAddress.Split('.').Last().ToLower();
                    picType = picType == "jpg" ? "jpeg" : picType;
                    ImagePart imagePart = null;
                    if (Enum.TryParse(picType, true, out ImagePartType imagePartType))
                    {
                        imagePart = main.AddImagePart(imagePartType);
                    }
                    var fs = System.IO.File.Open(picAddress, FileMode.Open);
                    imagePart?.FeedData(fs);
                    System.Drawing.Image image = System.Drawing.Image.FromStream(fs);
                    double width = image.Width;
                    double height = image.Height;
                    while (Math.Round(width / resolution * 2.54, 1) > initWidth || Math.Round(height / resolution * 2.54, 1) > initHeight)
                    {
                        width /= 2;
                        height /= 2;
                    }
                    Run run = new();
                    run.AddImageToBodyTableCell(main.GetIdOfPart(imagePart), (long)(width / resolution * 2.54 * 360000), (long)(height / resolution * 2.54 * 360000));
                    docBody.AppendChild(new Paragraph(run));
                    image.Dispose();
                    await fs.DisposeAsync();
                    docBody.Wrap(15);
                }
                else
                    docBody.Wrap(30);
            }
            else
                docBody.Wrap(30);
        }

    }
}
