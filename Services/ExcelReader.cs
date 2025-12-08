using OfficeOpenXml;
using System.Collections.Generic;

namespace juzi.task1.Services;

public class ExcelSheetData
{
    public string SheetName { get; set; } = "";
    public string Title { get; set; } = "";
    public List<Dictionary<string, object>> Data { get; set; } = new();
}

public class ExcelReader
{
    public static List<ExcelSheetData> ReadAllSheets(string filePath)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        var result = new List<ExcelSheetData>();
        
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            foreach (var worksheet in package.Workbook.Worksheets)
            {
                var sheetData = ReadSheet(worksheet);
                if (sheetData != null)
                {
                    result.Add(sheetData);
                }
            }
        }
        
        return result;
    }
    
    private static ExcelSheetData? ReadSheet(ExcelWorksheet worksheet)
    {
        var rowCount = worksheet.Dimension?.Rows ?? 0;
        var colCount = worksheet.Dimension?.Columns ?? 0;
        
        if (rowCount < 4 || colCount < 7)
            return null;
        
        var sheetData = new ExcelSheetData
        {
            SheetName = worksheet.Name
        };
        
        // 读取标题（第1行，C1单元格，可能是合并单元格）
        var titleCell = worksheet.Cells[1, 3]; // C1
        sheetData.Title = titleCell.Text.Trim();
        
        // 如果标题为空，尝试从合并单元格读取
        if (string.IsNullOrWhiteSpace(sheetData.Title))
        {
            var mergedCells = worksheet.MergedCells;
            foreach (var mergedRange in mergedCells)
            {
                // 检查合并单元格范围是否包含第1行第3列（C1）
                var range = worksheet.Cells[mergedRange];
                if (range.Start.Row <= 1 && range.End.Row >= 1 && 
                    range.Start.Column <= 3 && range.End.Column >= 3)
                {
                    sheetData.Title = range.Text.Trim();
                    break;
                }
            }
        }
        
        // 定义列结构（根据实际Excel格式）
        var columnHeaders = new List<string>
        {
            "项目",           // A列
            "办事处预算",     // B列
            "公司预算",       // C列
            "本月批复数",     // D列
            "上月批复数",     // E列
            "上月执行数",     // F列
            "备注"            // G列
        };
        
        // 从第4行开始读取数据（跳过标题和表头）
        for (int row = 4; row <= rowCount; row++)
        {
            var rowData = new Dictionary<string, object>();
            bool hasData = false;
            
            // 读取A列（项目名称）
            var projectName = worksheet.Cells[row, 1].Text?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(projectName))
            {
                continue; // 跳过空行
            }
            
            // 读取项目名称
            rowData[columnHeaders[0]] = projectName;
            hasData = true;
            
            // 读取B到G列的数据
            for (int col = 2; col <= 7; col++)
            {
                var cellValue = worksheet.Cells[row, col].Value;
                var headerIndex = col - 1;
                
                if (cellValue != null)
                {
                    // 处理数字格式
                    if (cellValue is double dblValue)
                    {
                        rowData[columnHeaders[headerIndex]] = dblValue;
                    }
                    else
                    {
                        rowData[columnHeaders[headerIndex]] = cellValue.ToString() ?? "";
                    }
                    hasData = true;
                }
                else
                {
                    rowData[columnHeaders[headerIndex]] = "";
                }
            }
            
            if (hasData)
            {
                sheetData.Data.Add(rowData);
            }
        }
        
        return sheetData;
    }
}

