# 方法1: 使用项目配置
dotnet publish -c Release

# 方法2: 命令行参数
dotnet publish -c Release -r win-x86 --self-contained true -p:PublishSingleFile=true -p:PublishTrimmed=true