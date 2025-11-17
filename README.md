# 企业微信文档批量下载工具

一个基于 Python + Selenium 的自动化工具，用于批量下载企业微信文档（表格和文档），支持目录遍历和断点续传。

## 功能特点

- ✅ **批量下载**：自动遍历目录树，批量下载企业微信文档和表格
- 📁 **目录遍历**：自动扫描根目录下所有包含 `data.json` 的子目录
- ⏭️ **断点续传**：自动跳过已下载的文件，支持中断后继续
- 🔄 **智能重试**：下载失败自动记录，方便后续重试
- 📊 **进度显示**：实时显示下载进度和统计信息
- 🔐 **多种登录方式**：支持 Cookie 注入或使用真实浏览器 Profile
- 💾 **自动重命名**：自动处理文件名冲突，生成编号文件名
- 📈 **详细日志**：完整的操作日志和结果报告

## 系统要求

- Python 3.7+
- Chrome 浏览器
- 稳定的网络连接

## 安装依赖

```bash
pip install undetected-chromedriver selenium
```

## 配置说明

### 1. 基本配置

在脚本开头的配置区修改以下参数：

```python
# 根目录配置 - 包含所有子目录的根目录
ROOT_DIRECTORY = "/path/to/your/root/directory"

# Cookie 文件路径（与脚本同目录）
cookie_file = "cookies.json"

# 浏览器配置
USE_REAL_PROFILE = False  # 是否使用真实浏览器 Profile
CHROME_PROFILE_PATH = ""  # Chrome Profile 路径
PROFILE_NAME = "Default"  # Profile 名称

headless = False  # 是否使用无头模式
```

### 2. 目录结构要求

工具会遍历 `ROOT_DIRECTORY` 下的所有子目录，查找包含 `data.json` 的目录并处理其中的文档。

期望的目录结构：
```
ROOT_DIRECTORY/
├── 项目A/
│   ├── data.json
│   └── (下载的文件会保存到这里)
├── 项目B/
│   ├── data.json
│   └── (下载的文件会保存到这里)
└── 子目录/
    └── 项目C/
        ├── data.json
        └── (下载的文件会保存到这里)
```

### 3. data.json 格式

每个目录下的 `data.json` 应包含以下结构：

```json
{
  "body": {
    "file_list": [
      {
        "name": "文档名称",
        "doc_url": "https://doc.weixin.qq.com/..."
      }
    ]
  }
}
```

### 4. 登录方式配置

#### 方式一：使用 Cookie（推荐用于自动化）

1. 手动登录企业微信文档
2. 使用浏览器开发者工具导出 Cookie
3. 保存为 `cookies.json`，支持两种格式：

**格式1：对象格式**
```json
{
  "cookie_name1": "value1",
  "cookie_name2": "value2"
}
```

**格式2：数组格式**
```json
[
  {
    "name": "cookie_name",
    "value": "cookie_value",
    "domain": ".weixin.qq.com",
    "path": "/",
    "secure": true
  }
]
```

#### 方式二：使用真实浏览器 Profile

```python
USE_REAL_PROFILE = True
CHROME_PROFILE_PATH = "/path/to/Chrome/User Data"
PROFILE_NAME = "Default"  # 或 "Profile 1" 等
```

## 使用方法

### 1. 准备工作

```bash
# 1. 克隆或下载脚本
# 2. 安装依赖
pip install undetected-chromedriver selenium

# 3. 修改配置
# 编辑脚本中的 ROOT_DIRECTORY 等配置

# 4. 准备 Cookie 或配置 Profile
# 将 cookies.json 放在脚本同目录下
```

### 2. 运行脚本

```bash
python download_wechat_docs.py
```

### 3. 运行过程

1. 脚本会自动扫描根目录下所有包含 `data.json` 的子目录
2. 显示找到的目录列表
3. 启动 Chrome 浏览器
4. 如果使用 Cookie 模式，会自动注入 Cookie
5. 遍历每个目录，下载其中的文档
6. 自动跳过已存在的文件
7. 生成详细的下载报告

## 输出说明

### 1. 控制台输出

```
================================================================================
🚀 企业微信文档批量下载工具（目录遍历版）
================================================================================
📁 根目录: /path/to/root
⏰ 开始时间: 2025-01-15 10:30:00
================================================================================

🔍 正在扫描目录...

✅ 找到 3 个包含data.json的目录:
   1. 项目A
   2. 项目B
   3. 子目录/项目C

================================================================================
📂 [1/3] 处理目录: 项目A
   路径: /path/to/root/项目A
================================================================================
📊 本目录共有 5 个文档待处理

📄 [1/5] 正在处理: 需求文档
   URL: https://doc.weixin.qq.com/...
✅ 下载完成: 需求文档.docx (2.5 MB)

...

📊 目录 [项目A] 处理完成
✅ 成功: 4
⏭️  跳过: 0
❌ 失败: 1
⏱️  耗时: 2分30秒
```

### 2. 结果文件

脚本会在根目录生成 `download_result_YYYYMMDD_HHMMSS.json`：

```json
{
  "start_time": "2025-01-15 10:30:00",
  "end_time": "2025-01-15 10:45:30",
  "total_time_seconds": 930,
  "total_success": 45,
  "total_failed": 2,
  "total_skipped": 3,
  "total_directories": 3,
  "directory_results": [
    {
      "directory": "项目A",
      "success": 15,
      "failed": 1,
      "skipped": 0,
      "status": "完成"
    }
  ]
}
```

### 3. 调试文件

如果下载失败，会在对应目录的 `debug/` 子目录下生成：
- HTML 文件：页面源码
- PNG 文件：页面截图

## 高级配置

### 超时设置

```python
PAGE_LOAD_TIMEOUT = 30  # 页面加载超时（秒）
WAIT_TIMEOUT = 15       # 元素等待超时（秒）
DOWNLOAD_TIMEOUT = 120  # 下载超时（秒）
```

### 等待时间

```python
MENU_WAIT = 2          # 菜单点击后等待（秒）
CLICK_WAIT = 1         # 按钮点击后等待（秒）
PAGE_STABLE_WAIT = 3   # 页面稳定等待（秒）
```

## 常见问题

### 1. 下载失败怎么办？

- 检查网络连接
- 确认 Cookie 或 Profile 是否有效
- 查看 `debug/` 目录下的调试文件
- 适当增加超时时间

### 2. 如何获取 Cookie？

1. 打开 Chrome 浏览器
2. 登录企业微信文档
3. 按 F12 打开开发者工具
4. 切换到 Application/存储 标签
5. 左侧选择 Cookies
6. 复制所需的 Cookie 值

### 3. 文件名冲突怎么办？

脚本会自动处理文件名冲突：
- 如果文件已存在，自动跳过
- 如果下载的文件名冲突，自动添加编号：`文件名(1).xlsx`

### 4. 如何只下载某些目录？

修改 `ROOT_DIRECTORY` 指向特定目录，或者在目录中创建/删除 `data.json` 文件来控制哪些目录被处理。

### 5. 中断后如何继续？

直接重新运行脚本即可，工具会自动跳过已下载的文件。

## 注意事项

⚠️ **重要提示**

1. 请确保有稳定的网络连接
2. 建议在网络状况良好时运行
3. 大量文档下载可能需要较长时间，请耐心等待
4. 定期检查 Cookie 或 Profile 的有效性
5. 建议先在小范围测试后再批量运行
6. 请遵守企业微信的使用条款和相关规定

## 性能优化建议

- 根据网络状况调整超时时间
- 使用真实浏览器 Profile 可以提高稳定性
- 分批处理大量文档，避免单次运行时间过长
- 定期清理下载目录中的临时文件

## 技术支持

如遇到问题，请检查：
1. Python 和依赖包版本
2. Chrome 浏览器版本
3. 网络连接状态
4. Cookie 或 Profile 有效性
5. 日志文件中的错误信息

## 许可证

本工具仅供学习和个人使用，请勿用于商业用途。

---

**版本**: 1.0  
**最后更新**: 2025-11
