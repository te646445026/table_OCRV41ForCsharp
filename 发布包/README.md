# 电梯限速器检测报告生成系统 (.NET Framework 4.7.2 版本)

这是一个从 .NET 8.0 迁移到 .NET Framework 4.7.2 的项目，用于通过腾讯云OCR服务识别电梯限速器检测表格，并自动生成Word格式的检测报告。

## 功能特点

- 使用腾讯云OCR服务识别表格数据
- 自动提取检测数据并填充到Word模板中
- 生成标准格式的检测记录和检测报告
- 支持节假日和工作日计算
- 完整的日志记录功能

## 系统要求

- Windows 操作系统
- .NET Framework 4.7.2
- Visual Studio 2019 或更高版本

## 迁移说明

本项目是从 .NET 8.0 迁移到 .NET Framework 4.7.2 的版本，保持了所有原有功能和逻辑不变。

## 安装和配置

### 1. 安装依赖包

项目使用以下NuGet包：

- Newtonsoft.Json (13.0.3)
- NPOI (2.7.1)
- Microsoft.Extensions.DependencyInjection (3.1.32)
- Microsoft.Extensions.Configuration (3.1.32)
- Microsoft.Extensions.Configuration.Json (3.1.32)
- Microsoft.Extensions.Configuration.FileExtensions (3.1.32)
- Nager.Date (2.8.3)
- System.Memory (4.5.5)
- System.Threading.Tasks.Extensions (4.5.4)
- System.Runtime.CompilerServices.Unsafe (4.5.3)
- System.Buffers (4.5.1)
- System.Numerics.Vectors (4.5.0)
- Microsoft.Extensions.DependencyInjection.Abstractions (3.1.32)
- Microsoft.Extensions.Configuration.Abstractions (3.1.32)
- Microsoft.Extensions.FileProviders.Abstractions (3.1.32)
- Microsoft.Extensions.FileProviders.Physical (3.1.32)
- Microsoft.Extensions.FileSystemGlobbing (3.1.32)
- Microsoft.Extensions.Primitives (3.1.32)

有多种方式安装依赖包：

#### 方式一：使用NuGet包管理器控制台（推荐）

在Visual Studio中打开"工具" -> "NuGet包管理器" -> "包管理器控制台"，然后运行以下命令：

```powershell
Install-Package Newtonsoft.Json -Version 13.0.3
Install-Package NPOI -Version 2.7.1
Install-Package Microsoft.Extensions.DependencyInjection -Version 3.1.32
Install-Package Microsoft.Extensions.Configuration -Version 3.1.32
Install-Package Microsoft.Extensions.Configuration.Json -Version 3.1.32
Install-Package Microsoft.Extensions.Configuration.FileExtensions -Version 3.1.32
Install-Package Nager.Date -Version 2.8.3
Install-Package System.Memory -Version 4.5.5
Install-Package System.Threading.Tasks.Extensions -Version 4.5.4
Install-Package System.Runtime.CompilerServices.Unsafe -Version 4.5.3
Install-Package System.Buffers -Version 4.5.1
Install-Package System.Numerics.Vectors -Version 4.5.0
Install-Package Microsoft.Extensions.DependencyInjection.Abstractions -Version 3.1.32
Install-Package Microsoft.Extensions.Configuration.Abstractions -Version 3.1.32
Install-Package Microsoft.Extensions.FileProviders.Abstractions -Version 3.1.32
Install-Package Microsoft.Extensions.FileProviders.Physical -Version 3.1.32
Install-Package Microsoft.Extensions.FileSystemGlobbing -Version 3.1.32
Install-Package Microsoft.Extensions.Primitives -Version 3.1.32
```

#### 方式二：使用PowerShell脚本

运行项目目录下的 `install-packages.ps1` 脚本自动安装所有依赖包。

#### 方式三：使用批处理脚本

运行项目目录下的 `install-packages.bat` 脚本自动安装所有依赖包。

### 解决常见的NuGet包引用问题

如果遇到"未能找到类型或命名空间"错误，请尝试以下解决方案：

1. **清理并重新生成解决方案**
   - 在Visual Studio中选择"生成" -> "清理解决方案"
   - 然后选择"生成" -> "重新生成解决方案"

2. **手动还原NuGet包**
   - 右键点击解决方案资源管理器中的解决方案
   - 选择"还原NuGet程序包"

3. **检查packages.config文件**
   - 确保文件包含所有必需的包
   - 如果缺少某些包，可以手动添加或使用包管理器UI添加

### 2. 配置文件

项目需要以下配置文件：

- `key.json` - 存储腾讯云API密钥
- `default.json` - 存储默认路径配置

首次运行时，程序会提示输入API密钥并自动生成 `key.json` 文件。

### 3. Word模板文件

确保以下Word模板文件存在于工作目录中：

- 限速器测试记录模板4.docx
- 限速器测试报告模板4.docx

## 编译项目

详细编译指南请查看 [COMPILE_GUIDE.md](COMPILE_GUIDE.md) 文件。

## 使用方法

1. 编译项目
2. 运行生成的可执行文件
3. 根据提示选择操作模式：
   - 上传图片进行OCR识别
   - 使用已生成的识别结果文件
4. 程序会自动生成Word格式的检测记录和检测报告

## 测试

项目包含以下测试类：

- HolidayServiceTest - 假期服务测试
- LoggerTest - 日志记录器测试

可以通过命令行参数运行测试：
- `test-holiday` - 运行假期服务测试
- `test-logger` - 运行日志记录器测试

示例：
```
table_OCRV41ForCsharp_net_framework.exe test-holiday
```

## 注意事项

1. 确保网络连接正常，以便调用腾讯云OCR服务
2. 确保Word模板文件路径正确
3. 确保有足够的磁盘空间存储生成的文件
4. 首次运行时需要配置API密钥

## 迁移状态

- [x] 第一阶段：项目结构迁移
- [x] 第二阶段：代码文件迁移
- [x] 第三阶段：依赖包处理
- [x] 第四阶段：配置文件处理
- [ ] 第五阶段：功能验证