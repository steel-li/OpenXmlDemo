using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using A = DocumentFormat.OpenXml.Drawing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OpenXmlDemo.Api.Help
{
    public static class OpenXmlHelper
    {

        /// <summary>
        /// 设置段落属性
        /// </summary>
        /// <param name="docBody">文档主体</param>
        /// <param name="text">文档段落文本</param>
        /// <param name="font">字体类型</param>
        /// <param name="fontSize">字体大小：以半磅为单位</param>
        /// <param name="color">字体颜色</param>
        /// <param name="justification">字体对齐方向：（0：左侧;2：居中）</param>
        /// <param name="isWrap">是否换行</param>
        /// <param name="wrapNum">换行数量</param>
        /// <param name="isBold">是否加粗</param>
        public static void SetParagraph(this Body docBody, string text, string font = "宋体", string fontSize = "36", string color = "#000000", int justification = 2, bool isWrap = true, int wrapNum = 1, bool isBold = false)
        {
            // 新增段落
            var para = docBody.AppendChild(new Paragraph());
            // 段落属性
            var paragraphProperties = para.AppendChild(new ParagraphProperties());
            paragraphProperties.Justification = paragraphProperties.AppendChild(new Justification() { Val = (JustificationValues)justification });

            var run = para.AppendChild(new Run());
            var runProperties = run.AppendChild(new RunProperties());
            run.AppendChild(new Text(text));
            runProperties.AppendChild(new RunFonts() { Ascii = font, HighAnsi = font, EastAsia = font });
            // 设置自动大小为18磅，以半磅为单位
            runProperties.AppendChild(new FontSize() { Val = fontSize });
            // 设置字体颜色
            runProperties.AppendChild(new Color() { Val = color });
            // 设置字体加粗
            runProperties.AppendChild(new Bold() { Val = new OnOffValue() { Value = isBold } });

            if (isWrap)
                docBody.Wrap(wrapNum, fontSize);

        }

        /// <summary>
        /// 初始化表格
        /// </summary>
        /// <param name="table"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="rowHeightType">0:自动高度(设置高度无效，需在设置上下左右间距);1:自定义高度(设置此模式后内容超出会被隐藏);</param>
        /// <param name="borderColor">4F81BD</param>
        /// <param name="tableWidth"></param>
        /// <param name="rowHeight"></param>
        /// <param name="leftMargion"></param>
        /// <param name="rightMargin"></param>
        /// <param name="topMargion"></param>
        /// <param name="bottomMargion"></param>
        /// <param name="cellAlignmentMethod">单元格对齐方式（0:上对齐 1:居中 2:下对齐）</param>
        public static void InitTable(this Table table, int row, int col, int rowHeightType = 0, string borderColor = "000000", string tableWidth = "5000", uint rowHeight = 600, string leftMargion = "100", string rightMargin = "100", string topMargion = "0", string bottomMargion = "0", int cellAlignmentMethod = 1)
        {
            // 设置表格边框
            var tabProps = table.AppendChild(new TableProperties()
            {
                TableBorders = new TableBorders()
                {
                    TopBorder = new TopBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4,
                        Color = borderColor
                    },
                    BottomBorder = new BottomBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4,
                        Color = borderColor
                    },
                    LeftBorder = new LeftBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4,
                        Color = borderColor
                    },
                    RightBorder = new RightBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4,
                        Color = borderColor
                    },
                    InsideHorizontalBorder = new InsideHorizontalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4,
                        Color = borderColor
                    },
                    InsideVerticalBorder = new InsideVerticalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 4,
                        Color = borderColor
                    }
                },
            });
            tabProps.AppendChild(new TableWidth { Width = tableWidth, Type = TableWidthUnitValues.Pct });
            // 初始化表格
            string width = (Convert.ToDouble(tableWidth) / col).ToString("0.0");
            for (int i = 0; i < row; i++)
            {
                TableRow tabRow = table.AppendChild(new TableRow());
                tabRow.AppendChild(new TableRowProperties(new TableRowHeight { Val = UInt32Value.FromUInt32(rowHeight), HeightType = (HeightRuleValues)rowHeightType }));
                for (int j = 0; j < col; j++)
                {
                    TableCell tableCell = tabRow.AppendChild(new TableCell());
                    var cellPara = tableCell.AppendChild(new Paragraph());
                    cellPara.AppendChild(new ParagraphProperties());
                    var tableCellProps = tableCell.AppendChild(new TableCellProperties());
                    // 设置单元格字体水平居中，垂直居中
                    tableCellProps.AppendChild(new TableCellVerticalAlignment { Val = (TableVerticalAlignmentValues)cellAlignmentMethod });
                    tableCellProps.AppendChild(new Justification { Val = JustificationValues.Center });
                    // 单元格宽度
                    tableCellProps.AppendChild(new TableCellWidth { Width = width, Type = TableWidthUnitValues.Pct });
                    // 设置单元格上下左右间距
                    tableCellProps.AppendChild(new TableCellMargin()
                    {
                        LeftMargin = new LeftMargin { Width = leftMargion, Type = TableWidthUnitValues.Dxa },
                        RightMargin = new RightMargin { Width = rightMargin, Type = TableWidthUnitValues.Dxa },
                        TopMargin = new TopMargin { Width = topMargion, Type = TableWidthUnitValues.Dxa },
                        BottomMargin = new BottomMargin { Width = bottomMargion, Type = TableWidthUnitValues.Dxa }
                    });
                }
            }
        }

        /// <summary>
        /// 插入图片px/37.8*36=cm
        /// </summary>
        /// <param name="run"></param>
        /// <param name="relationshipId"></param>
        /// <param name="width">图片宽度</param>
        /// <param name="height">高度</param>
        public static void AddImageToBodyTableCell(this Run run, string relationshipId, long width = 990000L, long height = 792000L)
        {
            var element =
                 new Drawing(
                     new Inline(
                         //new Extent() { Cx = 990000L, Cy = 792000L }, // 调节图片大小
                         new Extent() { Cx = width, Cy = height }, // 调节图片大小
                         new SimplePosition() { X = 0, Y = 0 },
                         new VerticalPosition() { RelativeFrom = VerticalRelativePositionValues.Paragraph },
                         new HorizontalPosition() { RelativeFrom = HorizontalRelativePositionValues.InsideMargin, PositionOffset = new PositionOffset("36") },
                         new EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DocProperties()
                         {
                             Id = 1U,
                             Name = "Picture 1"
                         },
                         new NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = 0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                       "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = width, Cy = height }), //与上面的对准
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = 0U,
                         DistanceFromBottom = 0U,
                         DistanceFromLeft = 0U,
                         DistanceFromRight = 0U,
                         EditId = "50D07946"
                     });
            run.AppendChild(element);
            // Append the reference to body, the element should be in a Run.
            //wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));

        }

        /// <summary>
        /// 换行
        /// </summary>
        /// <param name="docBody"></param>
        /// <param name="wrapNum"></param>
        /// <param name="fontSize"></param>
        public static void Wrap(this Body docBody, int wrapNum = 1, string fontSize = "20")
        {
            for (int i = 0; i < wrapNum; i++)
            {
                var para = docBody.AppendChild(new Paragraph());
                var run = para.AppendChild(new Run());
                var runProperties = run.AppendChild(new RunProperties());
                run.AppendChild(new Text(""));
                // 设置自动大小为18磅，以半磅为单位
                runProperties.AppendChild(new FontSize() { Val = fontSize });
            }
        }

        public static void CellWrap(this Run run, int wrapNum = 1)
        {
            for (int i = 0; i < wrapNum; i++)
            {
                run.AppendChild(new Break());
            }
        }

        /// <summary>
        /// 创建分节符
        /// </summary>
        /// <params>width</params>
        /// <params>height</params>
        /// <params>orient:1 横向;0:纵向</params>
        /// <returns></returns>
        public static void CreatePortraitSectionBreakParagraph(this Body docBody, uint width = 11906U, uint height = 16838U, int orient = 0)
        {
            Paragraph paragraph1 = new() { RsidParagraphAddition = "00052B73", RsidRunAdditionDefault = "00052B73" };

            ParagraphProperties paragraphProperties1 = new();

            SectionProperties sectionProperties1 = new() { RsidR = "00052B73" };

            sectionProperties1.AppendChild(new PageSize() { Width = width, Height = height, Orient = (PageOrientationValues)orient });
            sectionProperties1.AppendChild(new Columns() { Space = "425" });
            sectionProperties1.AppendChild(new DocGrid() { Type = DocGridValues.Lines, LinePitch = 312 });

            paragraphProperties1.Append(sectionProperties1);

            paragraph1.Append(paragraphProperties1);
            docBody.AddChild(paragraph1);
        }

    }

}
