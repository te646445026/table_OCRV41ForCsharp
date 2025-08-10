# 日志功能配置指南

## 概述

本项目已集成 NLog 日志框架，提供文件和控制台双重日志输出功能。

## 功能特性

### 1. 双重输出
- **文件日志**：保存到 `logs/app-{日期}.log` 文件
- **控制台日志**：实时显示在控制台窗口

### 2. 日志级别
- **Debug**：仅写入文件，用于详细调试信息
- **Info**：写入文件和控制台，用于一般信息
- **Warning**：写入文件和控制台，用于警告信息
- **Error**：写入文件和控制台，用于错误信息

### 3. 自动归档
- 每日自动创建新的日志文件
- 旧日志文件自动归档到 `logs/archive/` 目录
- 保留最近 30 天的日志文件

## 配置文件说明

### nlog.config

```xml
<?xml version="1.0" encoding="utf-8" ?>
<nlog xmlns="http://www.nlog-project.org/schemas/NLog.xsd"
      xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  
  <targets>
    <!-- 文件日志目标 -->
    <target xsi:type="File" name="fileTarget"
            fileName="logs/app-${shortdate}.log"
            layout="${longdate} [${level:uppercase=true}] ${logger} - ${message} ${exception:format=tostring}"
            archiveFileName="logs/archive/app-{#}.log"
            archiveEvery="Day"
            archiveNumbering="Rolling"
            maxArchiveFiles="30"
            encoding="utf-8" />
    
    <!-- 控制台日志目标 -->
    <target xsi:type="Console" name="consoleTarget"
            layout="${time} [${level:uppercase=true}] ${message} ${exception:format=ShortType}" />
  </targets>
  
  <rules>
    <!-- 所有日志级别 Info 及以上写入文件和控制台 -->
    <logger name="*" minlevel="Info" writeTo="fileTarget,consoleTarget" />
    
    <!-- Debug 级别只写入文件 -->
    <logger name="*" level="Debug" writeTo="fileTarget" final="true" />
  </rules>
  
</nlog>
```

## 使用方法

### 1. 基本日志记录

```csharp
// 获取日志记录器
var logger = serviceProvider.GetService<ILogger<Program>>();

// 记录不同级别的日志
logger.LogDebug("调试信息：变量值为 {value}", someValue);
logger.LogInformation("应用程序启动");
logger.LogWarning("检测到潜在问题：{issue}", issueDescription);
logger.LogError(exception, "处理过程中发生错误");
```

### 2. 结构化日志

```csharp
// 使用参数化消息
logger.LogInformation("用户 {UserId} 执行了操作 {Action}", userId, actionName);

// 记录异常信息
try
{
    // 业务逻辑
}
catch (Exception ex)
{
    logger.LogError(ex, "执行 {Operation} 时发生异常", operationName);
}
```

## 日志格式说明

### 文件日志格式
```
2024-01-15 14:30:25.123 [INFO] Program - 应用程序启动
2024-01-15 14:30:26.456 [ERROR] TencentOcrService - OCR 请求失败 System.HttpRequestException: 网络连接超时
   at TencentOcrService.ProcessAsync() in C:\...\TencentOcrService.cs:line 45
```

### 控制台日志格式
```
14:30:25 [INFO] 应用程序启动
14:30:26 [ERROR] OCR 请求失败 HttpRequestException
```

## 自定义配置

### 修改日志级别

在 `nlog.config` 中修改 `minlevel` 属性：

```xml
<!-- 只记录 Warning 及以上级别 -->
<logger name="*" minlevel="Warn" writeTo="fileTarget,consoleTarget" />
```

### 修改文件路径

```xml
<!-- 自定义日志文件路径 -->
<target xsi:type="File" name="fileTarget"
        fileName="D:/MyApp/logs/app-${shortdate}.log" />
```

### 添加更多输出目标

```xml
<!-- 添加数据库日志目标 -->
<target xsi:type="Database" name="databaseTarget"
        connectionString="..."
        commandText="INSERT INTO Logs(Date,Level,Logger,Message) VALUES(@date,@level,@logger,@message)">
  <parameter name="@date" layout="${date}" />
  <parameter name="@level" layout="${level}" />
  <parameter name="@logger" layout="${logger}" />
  <parameter name="@message" layout="${message}" />
</target>
```

## 故障排除

### 1. 日志文件未生成
- 检查应用程序是否有写入 `logs` 目录的权限
- 确认 `nlog.config` 文件已正确复制到输出目录
- 检查 `nlog-internal.log` 文件获取详细错误信息
- 确认 `nlog.config` 文件在正确位置（与可执行文件同目录）

### 2. 控制台无日志输出
- 检查日志级别设置是否正确
- 确认 NLog.Extensions.Logging 包已正确安装
- 验证 NLog 配置是否正确加载

### 3. 性能问题
- 考虑调整日志级别，减少 Debug 日志输出
- 启用异步日志记录：

```xml
<target xsi:type="AsyncWrapper" name="asyncFileTarget">
  <target xsi:type="File" fileName="logs/app-${shortdate}.log" />
</target>
```

### 4. 常见问题
- **控制台有日志但文件无日志**：检查文件路径权限和 `createDirs` 设置
- **中文乱码**：确认 `encoding="utf-8"` 设置正确
- **日志重复**：检查是否有多个日志规则匹配同一日志

## 最佳实践

1. **使用结构化日志**：使用参数化消息而不是字符串拼接
2. **合理设置日志级别**：生产环境建议使用 Info 级别
3. **记录关键操作**：OCR 请求、文件处理、异常等
4. **避免敏感信息**：不要记录密码、密钥等敏感数据
5. **定期清理日志**：配置合适的归档策略，避免磁盘空间不足

## 升级说明

本次升级从纯控制台日志改为 NLog 双重输出：

- ✅ 保持原有的控制台输出
- ✅ 新增文件持久化存储
- ✅ 支持日志归档和清理
- ✅ 提供更丰富的配置选项
- ✅ 完全兼容现有代码