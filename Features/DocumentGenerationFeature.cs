using juzi.task1.Services;

namespace juzi.task1.Features;

public class DocumentGenerationFeature
{
    public static void Execute()
    {
        Console.WriteLine("\n=== 日常费用预算财务文档生成 ===");
        Console.WriteLine();
        
        // 获取月份输入
        int month = GetMonthInput();
        int year = DateTime.Now.Year;
        
        try
        {
            string excelPath = Path.Combine("data", "办事处日常费用预算财务.xlsx");
            string templatePath = Path.Combine("data", "日常费用预算财务.docx");
            string outputDir = "output";
            
            // 检查模板文件（尝试 .docx 和 .doc）
            if (!File.Exists(templatePath))
            {
                templatePath = Path.Combine("data", "日常费用预算财务.docx");
            }
            
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"❌ 错误: 找不到 Excel 文件: {excelPath}");
                return;
            }
            
            if (!File.Exists(templatePath))
            {
                Console.WriteLine($"❌ 错误: 找不到模板文件: {templatePath}");
                Console.WriteLine("   请确保模板文件存在: data/日常费用预算财务.docx 或 data/日常费用预算财务.docx");
                return;
            }
            
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }
            
            Console.WriteLine($"📄 正在读取 Excel 文件: {excelPath}");
            var sheetDataList = ExcelReader.ReadAllSheets(excelPath);
            
            if (sheetDataList.Count == 0)
            {
                Console.WriteLine("⚠️  警告: Excel 文件中没有找到有效的数据表");
                return;
            }
            
            Console.WriteLine($"✅ 发现 {sheetDataList.Count} 个工作表");
            Console.WriteLine($"📋 使用模板文件: {templatePath}");
            Console.WriteLine();
            
            int successCount = 0;
            int failCount = 0;
            
            // 为每个 sheet 生成对应的 Word 文档
            foreach (var sheetData in sheetDataList)
            {
                try
                {
                    // 使用 sheet 名称作为文件名
                    string fileName = $"{sheetData.SheetName}.docx";
                    string outputPath = Path.Combine(outputDir, fileName);
                    
                    Console.WriteLine($"📝 正在处理工作表: {sheetData.SheetName}");
                    Console.WriteLine($"   标题: {sheetData.Title}");
                    Console.WriteLine($"   数据行数: {sheetData.Data.Count}");
                    Console.WriteLine($"   月份: {year}年{month}月");
                    
                    WordTemplateProcessor.GenerateFromTemplate(templatePath, outputPath, sheetData, year, month);
                    Console.WriteLine($"   ✅ 已生成: {outputPath}");
                    Console.WriteLine();
                    
                    successCount++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"   ❌ 生成失败: {ex.Message}");
                    Console.WriteLine();
                    failCount++;
                }
            }
            
            Console.WriteLine("=====================================");
            Console.WriteLine($"✅ 成功生成: {successCount} 个文档");
            if (failCount > 0)
            {
                Console.WriteLine($"❌ 失败: {failCount} 个文档");
            }
            Console.WriteLine($"📁 输出目录: {Path.GetFullPath(outputDir)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"❌ 错误: {ex.Message}");
            Console.WriteLine($"   堆栈跟踪: {ex.StackTrace}");
        }
    }
    
    private static int GetMonthInput()
    {
        int currentMonth = DateTime.Now.Month;
        
        Console.Write($"请输入月份 (1-12，直接回车使用当前月份 {currentMonth}): ");
        string? input = Console.ReadLine();
        
        if (string.IsNullOrWhiteSpace(input))
        {
            Console.WriteLine($"✓ 使用当前月份: {currentMonth}");
            return currentMonth;
        }
        
        if (int.TryParse(input.Trim(), out int month))
        {
            if (month >= 1 && month <= 12)
            {
                Console.WriteLine($"✓ 使用月份: {month}");
                return month;
            }
            else
            {
                Console.WriteLine($"⚠️  月份无效，使用当前月份: {currentMonth}");
                return currentMonth;
            }
        }
        else
        {
            Console.WriteLine($"⚠️  输入格式错误，使用当前月份: {currentMonth}");
            return currentMonth;
        }
    }
}

