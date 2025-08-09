# 项目编译指导

本文档详细说明如何在Visual Studio中编译这个.NET Framework 4.7.2项目。

## 编译前准备

### 1. 环境要求
- Visual Studio 2019 或更高版本
- .NET Framework 4.7.2 开发工具包
- NuGet 包管理器（通常随Visual Studio一起安装）

### 2. 安装依赖包
在编译项目之前，需要先安装所有必需的NuGet包。

#### 方法一：使用NuGet包管理器UI（推荐）
1. 在Visual Studio中打开项目
2. 右键点击解决方案，选择"还原NuGet程序包"
   或者
   转到"工具" -> "NuGet包管理器" -> "管理解决方案的NuGet程序包"
3. 点击"还原"按钮安装所有包

#### 方法二：使用包管理器控制台
1. 在Visual Studio中打开"工具" -> "NuGet包管理器" -> "包管理器控制台"
2. 确保"默认项目"下拉菜单中选择了当前项目
3. 运行以下命令：

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

#### 方法三：使用PowerShell脚本
1. 在项目目录中找到并运行 `install-packages.ps1` 文件
2. 等待所有包安装完成

#### 方法四：使用批处理脚本
1. 在项目目录中找到并运行 `install-packages.bat` 文件
2. 等待所有包安装完成

## 解决常见的NuGet包引用问题

如果遇到"未能找到类型或命名空间"错误，请尝试以下解决方案：

### 1. 清理并重新生成解决方案
- 在Visual Studio中选择"生成" -> "清理解决方案"
- 然后选择"生成" -> "重新生成解决方案"

### 2. 手动还原NuGet包
- 右键点击解决方案资源管理器中的解决方案
- 选择"还原NuGet程序包"

### 3. 检查包管理器控制台输出
- 打开"工具" -> "NuGet包管理器" -> "包管理器控制台"
- 查看是否有任何错误信息

### 4. 检查packages.config文件
- 确保 [packages.config](file:///D:/Csharp/OCR/table_OCRV41ForCsharp/table_OCRV41ForCsharp_net_framework/packages.config) 文件包含所有必需的包
- 如果缺少某些包，可以手动添加或使用包管理器UI添加

## 编译步骤

### 1. 打开项目
- 启动Visual Studio
- 选择"文件" -> "打开" -> "项目/解决方案"
- 浏览到项目目录，选择 `table_OCRV41ForCsharp_net_framework.csproj` 文件

### 2. 检查项目配置
- 确保项目目标框架是 .NET Framework 4.7.2
- 检查构建配置是否为 Debug 或 Release

### 3. 还原NuGet包
- 在解决方案资源管理器中右键点击解决方案
- 选择"还原NuGet程序包"

### 4. 编译项目
有两种方式编译项目：

#### 方法一：使用菜单
1. 选择"生成" -> "生成解决方案"
2. 等待编译完成
3. 查看"输出"窗口中的结果

#### 方法二：使用快捷键
- 按 `Ctrl + Shift + B` 快捷键

### 5. 检查编译结果
- 如果编译成功，将在输出目录（bin\Debug\ 或 bin\Release\）中生成可执行文件
- 如果编译失败，请检查错误信息并解决相应问题

## 常见问题及解决方案

### 1. 缺少NuGet包
**问题**: 编译时出现类似 "找不到类型或命名空间" 的错误
**解决方案**: 
- 确保已正确还原所有NuGet包
- 检查 packages.config 文件中的包是否都已安装
- 尝试清理并重新生成解决方案

### 2. 程序集引用问题
**问题**: 出现 "无法加载文件或程序集" 错误
**解决方案**:
- 检查 App.config 中的绑定重定向配置是否正确
- 确保所有依赖包版本与绑定重定向中的版本匹配

### 3. .NET Framework版本问题
**问题**: 出现与框架版本相关的编译错误
**解决方案**:
- 确保已安装 .NET Framework 4.7.2 开发工具包
- 在项目属性中确认目标框架设置正确

### 4. Newtonsoft.Json相关错误
**问题**: 出现 "未能找到类型或命名空间名 Newtonsoft" 错误
**解决方案**:
- 确保已安装 Newtonsoft.Json 13.0.3 包
- 检查文件顶部是否包含 `using Newtonsoft.Json;` 或 `using Newtonsoft.Json.Linq;` 引用

### 5. NPOI相关错误
**问题**: 出现 "未能找到类型或命名空间名 NPOI" 错误
**解决方案**:
- 确保已安装 NPOI 2.7.1 包
- 检查文件顶部是否包含 `using NPOI.XWPF.UserModel;` 引用

### 6. Microsoft.Extensions相关错误
**问题**: 出现 "命名空间 Microsoft 中不存在类型或命名空间名 Extensions" 错误
**解决方案**:
- 确保已安装 Microsoft.Extensions.DependencyInjection 3.1.32 包
- 检查文件顶部是否包含 `using Microsoft.Extensions.DependencyInjection;` 引用

## 运行程序

编译成功后，可以通过以下方式运行程序：

### 1. 在Visual Studio中运行
- 按 `F5` 开始调试
- 或按 `Ctrl + F5` 开始执行（不调试）

### 2. 直接运行可执行文件
- 导航到输出目录（bin\Debug\ 或 bin\Release\）
- 双击可执行文件运行

### 3. 通过命令行运行
```cmd
cd bin\Debug\
table_OCRV41ForCsharp_net_framework.exe
```

## 运行测试

项目包含两个测试类，可以通过命令行参数运行：

```cmd
# 运行假期服务测试
table_OCRV41ForCsharp_net_framework.exe test-holiday

# 运行日志记录器测试
table_OCRV41ForCsharp_net_framework.exe test-logger
```

## 需要的外部文件

确保以下文件在运行时可用：

1. `key.json` - 包含腾讯云API密钥（首次运行时会自动生成）
2. `default.json` - 包含默认路径配置
3. Word模板文件：
   - 限速器测试记录模板4.docx
   - 限速器测试报告模板4.docx

这些文件应该放在与可执行文件相同的目录中。