using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Text;
using System.Text.RegularExpressions;

namespace juzi.task1.Services;

public class WordTemplateProcessor
{
    public static void GenerateFromTemplate(
        string templatePath, 
        string outputPath, 
        ExcelSheetData sheetData)
    {
        if (!File.Exists(templatePath))
        {
            throw new FileNotFoundException($"模板文件不存在: {templatePath}");
        }
        
        // 复制模板文件（如果是 .doc 格式，需要转换为 .docx）
        string tempOutputPath = outputPath;
        if (templatePath.EndsWith(".doc", StringComparison.OrdinalIgnoreCase) && 
            !outputPath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
        {
            tempOutputPath = outputPath.Replace(".doc", ".docx", StringComparison.OrdinalIgnoreCase);
        }
        
        // 如果模板是 .docx，直接复制；如果是 .doc，需要转换
        if (templatePath.EndsWith(".docx", StringComparison.OrdinalIgnoreCase))
        {
            File.Copy(templatePath, tempOutputPath, true);
        }
        else
        {
            // .doc 格式需要先转换为 .docx（这里假设用户会提供 .docx 格式的模板）
            // 如果确实是 .doc，可能需要使用其他库如 Aspose.Words 或 Microsoft.Office.Interop.Word
            throw new NotSupportedException("当前不支持 .doc 格式模板，请使用 .docx 格式");
        }
        
        // 打开并修改文档
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(tempOutputPath, true))
        {
            MainDocumentPart? mainPart = wordDoc.MainDocumentPart;
            if (mainPart == null)
            {
                throw new InvalidOperationException("无法读取 Word 文档主部分");
            }
            
            Body? body = mainPart.Document?.Body;
            if (body == null)
            {
                throw new InvalidOperationException("无法读取 Word 文档正文");
            }
            
            // 提取费用数据（使用"本月批复数"列）
            var expenseItems = new List<(string Name, double Amount)>();
            double totalAmount = 0;
            
            foreach (var row in sheetData.Data)
            {
                if (row.ContainsKey("项目") && row.ContainsKey("本月批复数"))
                {
                    string projectName = row["项目"]?.ToString() ?? "";
                    object? approvedValue = row["本月批复数"];
                    
                    // 跳过"合计"行和空行
                    if (string.IsNullOrWhiteSpace(projectName) || projectName == "合计" || projectName == "其他")
                        continue;
                    
                    double amount = 0;
                    if (approvedValue is double dbl)
                    {
                        amount = dbl;
                    }
                    else if (approvedValue != null && double.TryParse(approvedValue.ToString(), out double parsed))
                    {
                        amount = parsed;
                    }
                    
                    if (amount > 0)
                    {
                        expenseItems.Add((projectName, amount));
                        totalAmount += amount;
                    }
                }
            }
            
            // 替换文档内容
            ReplaceTextInDocument(body, "天河办事处", $"{sheetData.SheetName}办事处");
            
            // 替换费用列表
            ReplaceExpenseList(body, expenseItems, totalAmount);
            
            // 保存文档
            mainPart.Document?.Save();
        }
        
        // 如果输出路径不同，重命名文件
        if (tempOutputPath != outputPath && File.Exists(tempOutputPath))
        {
            if (File.Exists(outputPath))
                File.Delete(outputPath);
            File.Move(tempOutputPath, outputPath);
        }
    }
    
    private static void ReplaceTextInDocument(Body body, string oldText, string newText)
    {
        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            foreach (var run in paragraph.Descendants<Run>())
            {
                foreach (var text in run.Descendants<Text>())
                {
                    if (text.Text.Contains(oldText))
                    {
                        text.Text = text.Text.Replace(oldText, newText);
                    }
                }
            }
        }
    }
    
    private static void ReplaceExpenseList(Body body, List<(string Name, double Amount)> expenseItems, double totalAmount)
    {
        // 查找费用列表的开始位置（"批复如下："之后）
        bool foundStart = false;
        Paragraph? startPara = null;
        int startIndex = -1;
        
        // 手动查找段落索引
        int paraIndex = 0;
        foreach (var element in body.ChildElements)
        {
            if (element is Paragraph para)
            {
                string paraText = GetParagraphText(para);
                
                if (paraText.Contains("批复如下："))
                {
                    foundStart = true;
                    startPara = para;
                    startIndex = paraIndex;
                    break;
                }
                paraIndex++;
            }
        }
        
        if (!foundStart || startPara == null || startIndex < 0)
        {
            throw new InvalidOperationException("无法找到费用列表的开始位置");
        }
        
        // 查找费用列表的结束位置（"合计"或"请你处严格"之前）
        int endIndex = -1;
        bool foundEnd = false;
        
        for (int i = startIndex + 1; i < body.ChildElements.Count; i++)
        {
            var element = body.ChildElements[i];
            if (element is Paragraph para)
            {
                string paraText = GetParagraphText(para);
                
                if (paraText.Contains("合计") && paraText.Contains("元") && 
                    (paraText.Contains("请你处严格") || paraText.Contains("严格按费用明细")))
                {
                    endIndex = i;
                    foundEnd = true;
                    break;
                }
            }
        }
        
        if (!foundEnd)
        {
            // 如果没找到明确的结束位置，查找包含"合计"和"元"的段落
            for (int i = startIndex + 1; i < body.ChildElements.Count; i++)
            {
                var element = body.ChildElements[i];
                if (element is Paragraph para)
                {
                    string paraText = GetParagraphText(para);
                    if (paraText.Contains("合计") && paraText.Contains("元"))
                    {
                        endIndex = i;
                        foundEnd = true;
                        break;
                    }
                }
            }
        }
        
        if (!foundEnd)
        {
            throw new InvalidOperationException("无法找到费用列表的结束位置");
        }
        
        // 删除旧的费用列表段落
        for (int i = endIndex; i > startIndex; i--)
        {
            var element = body.ChildElements[i];
            if (element is Paragraph)
            {
                body.RemoveChild(element);
            }
        }
        
        // 插入新的费用列表
        int insertIndex = startIndex + 1;
        
        foreach (var (name, amount) in expenseItems)
        {
            Paragraph para = new Paragraph();
            
            // 设置段落缩进（前面空2格）
            ParagraphProperties paraProps = new ParagraphProperties();
            Indentation indent = new Indentation() { FirstLine = "720" }; // 720 = 2个中文字符的缩进
            paraProps.Append(indent);
            para.Append(paraProps);
            
            // 费用名称（加粗、楷体GB2312、4号字）
            Run nameRun = new Run();
            RunProperties nameRunProps = new RunProperties();
            nameRunProps.Bold = new Bold();
            // 设置字体为楷体GB2312
            nameRunProps.RunFonts = new RunFonts() { EastAsia = "楷体_GB2312", Ascii = "KaiTi_GB2312", HighAnsi = "KaiTi_GB2312" };
            // 设置字号为4号（14磅 = 28 half-points）
            nameRunProps.FontSize = new FontSize() { Val = "28" };
            nameRun.Append(nameRunProps);
            Text nameText = new Text(name);
            nameRun.Append(nameText);
            para.Append(nameRun);
            
            // 制表符和金额（楷体GB2312、4号字、加粗）
            Run amountRun = new Run();
            RunProperties amountRunProps = new RunProperties();
            amountRunProps.Bold = new Bold();
            // 设置字体为楷体GB2312
            amountRunProps.RunFonts = new RunFonts() { EastAsia = "楷体_GB2312", Ascii = "KaiTi_GB2312", HighAnsi = "KaiTi_GB2312" };
            // 设置字号为4号（14磅 = 28 half-points）
            amountRunProps.FontSize = new FontSize() { Val = "28" };
            amountRun.Append(amountRunProps);
            Text tabText = new Text("\t");
            amountRun.Append(tabText);
            Text amountText = new Text(amount.ToString());
            amountRun.Append(amountText);
            para.Append(amountRun);
            
            body.InsertAt(para, insertIndex++);
        }
        
        // 插入合计行（合计行也需要缩进）
        Paragraph totalPara = new Paragraph();
        ParagraphProperties totalParaProps = new ParagraphProperties();
        Indentation totalIndent = new Indentation() { FirstLine = "720" };
        totalParaProps.Append(totalIndent);
        totalPara.Append(totalParaProps);
        
        // "合计"加粗、楷体GB2312、4号字
        Run totalLabelRun = new Run();
        RunProperties totalLabelRunProps = new RunProperties();
        totalLabelRunProps.Bold = new Bold();
        // 设置字体为楷体GB2312
        totalLabelRunProps.RunFonts = new RunFonts() { EastAsia = "楷体_GB2312", Ascii = "KaiTi_GB2312", HighAnsi = "KaiTi_GB2312" };
        // 设置字号为4号（14磅 = 28 half-points）
        totalLabelRunProps.FontSize = new FontSize() { Val = "28" };
        totalLabelRun.Append(totalLabelRunProps);
        Text totalLabelText = new Text("合计");
        totalLabelRun.Append(totalLabelText);
        totalPara.Append(totalLabelRun);
        
        // 金额部分（楷体GB2312、4号字、加粗）
        Run totalAmountRun = new Run();
        RunProperties totalAmountRunProps = new RunProperties();
        totalAmountRunProps.Bold = new Bold();
        totalAmountRunProps.RunFonts = new RunFonts() { EastAsia = "楷体_GB2312", Ascii = "KaiTi_GB2312", HighAnsi = "KaiTi_GB2312" };
        totalAmountRunProps.FontSize = new FontSize() { Val = "28" };
        totalAmountRun.Append(totalAmountRunProps);
        Text totalAmountText = new Text($"\t{totalAmount}元");
        totalAmountRun.Append(totalAmountText);
        totalPara.Append(totalAmountRun);
        
        // 后续说明文字（正常格式）
        Run totalContentRun = new Run();
        Text totalContentText = new Text("，请你处严格按费用明细开支，并按财务制度规定，务必于12月15日前将本月相关合法单据寄到财务部核销，逾期不予报销。");
        totalContentRun.Append(totalContentText);
        totalPara.Append(totalContentRun);
        
        body.InsertAt(totalPara, insertIndex);
    }
    
    private static string GetParagraphText(Paragraph paragraph)
    {
        var textBuilder = new StringBuilder();
        foreach (var run in paragraph.Descendants<Run>())
        {
            foreach (var text in run.Descendants<Text>())
            {
                textBuilder.Append(text.Text);
            }
        }
        return textBuilder.ToString();
    }
}

