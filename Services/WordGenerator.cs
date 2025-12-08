using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;

namespace juzi.task1.Services;

public class WordGenerator
{
    public static void GenerateWordDocument(string outputPath, Dictionary<string, string> data)
    {
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(
            outputPath, WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());
            
            // 添加标题
            if (data.ContainsKey("标题"))
            {
                Paragraph titlePara = body.AppendChild(new Paragraph());
                Run titleRun = titlePara.AppendChild(new Run());
                titleRun.AppendChild(new Text(data["标题"]));
                
                RunProperties titleProps = titleRun.AppendChild(new RunProperties());
                titleProps.FontSize = new FontSize() { Val = "32" };
                titleProps.Bold = new Bold();
                
                ParagraphProperties paraProps = titlePara.AppendChild(new ParagraphProperties());
                paraProps.Justification = new Justification() { Val = JustificationValues.Center };
            }
            
            // 添加空行
            body.AppendChild(new Paragraph());
            
            // 添加内容
            foreach (var kvp in data)
            {
                if (kvp.Key == "标题")
                    continue;
                    
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                
                string text = $"{kvp.Key}: {kvp.Value}";
                run.AppendChild(new Text(text));
                
                body.AppendChild(new Paragraph());
            }
            
            mainPart.Document.Save();
        }
    }
    
    public static void GenerateWordFromSheetData(string outputPath, ExcelSheetData sheetData)
    {
        if (sheetData == null || sheetData.Data.Count == 0)
        {
            throw new ArgumentException("Sheet 数据为空");
        }
        
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(
            outputPath, WordprocessingDocumentType.Document))
        {
            MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());
            
            // 添加标题
            Paragraph titlePara = body.AppendChild(new Paragraph());
            Run titleRun = titlePara.AppendChild(new Run());
            string title = string.IsNullOrWhiteSpace(sheetData.Title) 
                ? $"{sheetData.SheetName}日常费用预算财务" 
                : sheetData.Title;
            titleRun.AppendChild(new Text(title));
            
            RunProperties titleProps = titleRun.AppendChild(new RunProperties());
            titleProps.FontSize = new FontSize() { Val = "32" };
            titleProps.Bold = new Bold();
            
            ParagraphProperties titleParaProps = titlePara.AppendChild(new ParagraphProperties());
            titleParaProps.Justification = new Justification() { Val = JustificationValues.Center };
            titleParaProps.SpacingBetweenLines = new SpacingBetweenLines() { After = "200" };
            
            body.AppendChild(new Paragraph());
            
            // 添加表格
            Table table = body.AppendChild(new Table());
            
            // 表格属性
            TableProperties tableProps = table.AppendChild(new TableProperties());
            tableProps.TableWidth = new TableWidth() { Width = "5000", Type = TableWidthUnitValues.Pct };
            tableProps.TableBorders = new TableBorders(
                new TopBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new BottomBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new LeftBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new RightBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new InsideHorizontalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 },
                new InsideVerticalBorder() { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 4 }
            );
            
            // 添加多级表头
            // 第一级表头（合并单元格）
            TableRow headerRow1 = table.AppendChild(new TableRow());
            
            // A列：项目（需要跨两行）
            TableCell projectCell = AddHeaderCell(headerRow1, "项目", 1);
            projectCell.AppendChild(new TableCellProperties(
                new VerticalMerge() { Val = MergedCellValues.Restart }
            ));
            
            // B列：本月费用预算（合并B-C列，水平合并）
            TableCell budgetCell = AddHeaderCell(headerRow1, "本月费用预算", 2);
            
            // D列：本月批复数（需要跨两行）
            TableCell approvedCell = AddHeaderCell(headerRow1, "本月批复数", 1);
            approvedCell.AppendChild(new TableCellProperties(
                new VerticalMerge() { Val = MergedCellValues.Restart }
            ));
            
            // E列：上月费用批复数/执行数（合并E-F列，水平合并）
            TableCell lastMonthCell = AddHeaderCell(headerRow1, "上月费用批复数/执行数", 2);
            
            // G列：备注（需要跨两行）
            TableCell remarkCell = AddHeaderCell(headerRow1, "备注", 1);
            remarkCell.AppendChild(new TableCellProperties(
                new VerticalMerge() { Val = MergedCellValues.Restart }
            ));
            
            // 第二级表头
            TableRow headerRow2 = table.AppendChild(new TableRow());
            
            // A列：继续合并（垂直）
            TableCell projectCell2 = AddHeaderCell(headerRow2, "", 1);
            projectCell2.AppendChild(new TableCellProperties(
                new VerticalMerge() { Val = MergedCellValues.Continue }
            ));
            
            // B列：办事处预算
            AddHeaderCell(headerRow2, "办事处预算", 1);
            
            // C列：公司预算
            AddHeaderCell(headerRow2, "公司预算", 1);
            
            // D列：继续合并（垂直）
            TableCell approvedCell2 = AddHeaderCell(headerRow2, "", 1);
            approvedCell2.AppendChild(new TableCellProperties(
                new VerticalMerge() { Val = MergedCellValues.Continue }
            ));
            
            // E列：批复数
            AddHeaderCell(headerRow2, "批复数", 1);
            
            // F列：执行数
            AddHeaderCell(headerRow2, "执行数", 1);
            
            // G列：继续合并（垂直）
            TableCell remarkCell2 = AddHeaderCell(headerRow2, "", 1);
            remarkCell2.AppendChild(new TableCellProperties(
                new VerticalMerge() { Val = MergedCellValues.Continue }
            ));
            
            // 添加数据行
            var headers = new List<string> { "项目", "办事处预算", "公司预算", "本月批复数", "上月批复数", "上月执行数", "备注" };
            
            foreach (var rowData in sheetData.Data)
            {
                TableRow row = table.AppendChild(new TableRow());
                
                foreach (var header in headers)
                {
                    TableCell cell = row.AppendChild(new TableCell());
                    Paragraph para = cell.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());
                    
                    object? value = rowData.ContainsKey(header) ? rowData[header] : "";
                    string displayValue = "";
                    
                    if (value is double dblValue)
                    {
                        displayValue = dblValue == 0 ? "" : dblValue.ToString("F0");
                    }
                    else
                    {
                        displayValue = value?.ToString() ?? "";
                    }
                    
                    run.AppendChild(new Text(displayValue));
                    
                    // 合计行加粗
                    if (rowData.ContainsKey("项目") && rowData["项目"]?.ToString() == "合计")
                    {
                        RunProperties runProps = run.AppendChild(new RunProperties());
                        runProps.Bold = new Bold();
                    }
                    
                    ParagraphProperties paraProps = para.AppendChild(new ParagraphProperties());
                    if (header == "项目")
                    {
                        paraProps.Justification = new Justification() { Val = JustificationValues.Left };
                    }
                    else
                    {
                        paraProps.Justification = new Justification() { Val = JustificationValues.Right };
                    }
                }
            }
            
            mainPart.Document.Save();
        }
    }
    
    private static TableCell AddHeaderCell(TableRow row, string text, int colspan)
    {
        TableCell cell = row.AppendChild(new TableCell());
        Paragraph para = cell.AppendChild(new Paragraph());
        Run run = para.AppendChild(new Run());
        run.AppendChild(new Text(text));
        
        RunProperties runProps = run.AppendChild(new RunProperties());
        runProps.Bold = new Bold();
        
        ParagraphProperties paraProps = para.AppendChild(new ParagraphProperties());
        paraProps.Justification = new Justification() { Val = JustificationValues.Center };
        
        if (colspan > 1)
        {
            cell.AppendChild(new TableCellProperties(new GridSpan() { Val = colspan }));
        }
        
        return cell;
    }
    
}

