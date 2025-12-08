using juzi.task1.Features;

namespace juzi.task1;

class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("=====================================");
        Console.WriteLine("   日常费用预算财务文档生成工具");
        Console.WriteLine("=====================================");
        
        while (true)
        {
            ShowMenu();
            
            string? choice = Console.ReadLine();
            Console.WriteLine();
            
            switch (choice?.Trim())
            {
                case "1":
                    DocumentGenerationFeature.Execute();
                    break;
                    
                case "0":
                    Console.WriteLine("感谢使用，再见！");
                    return;
                    
                default:
                    Console.WriteLine("❌ 无效的选择，请重新输入。");
                    break;
            }
            
            Console.WriteLine("\n按任意键继续...");
            Console.ReadKey();
            Console.Clear();
        }
    }
    
    static void ShowMenu()
    {
        Console.WriteLine("\n请选择功能：");
        Console.WriteLine("  1. 生成日常费用预算财务文档");
        Console.WriteLine("  0. 退出");
        Console.Write("\n请输入选项 (0-1): ");
    }
}
