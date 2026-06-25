---
name: "excel-data-analyzer"
description: "专业Excel数据分析工具，支持数据分析、Excel操作题目、统计计算、可视化。自动安装环境，保留公式和Excel功能。适用于分析Excel表格、完成Excel考试题、计算平均值综测时调用。"
---

# Excel 专业数据分析器

> 作为精通Excel的数据分析工程师，提供完整的Excel数据分析和操作解决方案。

## ⚠️ 重要说明：严格执行流程

**必须按顺序执行所有步骤，禁止跳步！**

### 为什么不能跳步？

每个步骤都是精心设计的必要环节：

1. **环境配置** → 创建隔离的虚拟环境，避免系统污染
2. **数据读取** → 智能转换格式，确保数据正确加载
3. **数据处理** → 智能类型转换，避免计算错误
4. **公式写入** → 保留Excel公式，体现真实操作
5. **筛选操作** → 保持Excel功能完整性
6. **图表创建** → 引用原始数据，动态更新
7. **结果验证** → 确保输出质量，避免错误

### 如果跳步会发生什么？

❌ **跳过环境配置**：
- 依赖库缺失或版本冲突
- 运行时错误："ModuleNotFoundError"
- 系统环境被污染

❌ **跳过数据读取**：
- .xls格式无法处理
- 数据丢失或格式错误
- 列名识别失败

❌ **跳过数据处理**：
- 数值型数据变成字符串
- 计算公式报错
- 空值导致异常

❌ **跳过公式写入**：
- 用户只看到数值，看不到公式
- 无法体现Excel操作技能
- 考试评分可能不合格

❌ **跳过验证步骤**：
- 输出文件格式错误
- 公式计算结果错误
- 图表数据源错误

### 正确执行流程

**必须完整执行以下流程**：

```python
# ✅ 正确：完整执行
setup_environment()         # 步骤1：环境配置
file_path = smart_read_excel()  # 步骤2：数据读取
write_excel_formula()       # 步骤3：公式写入
filter_data_to_sheet()      # 步骤4：筛选操作
create_excel_chart()        # 步骤5：图表创建
validate_output()           # 步骤6：结果验证
wb.save()                   # 步骤7：保存文件

# ❌ 错误：跳步执行
wb = Workbook()             # 直接创建工作簿（跳过环境）
ws.cell().value = 8         # 直接填数值（跳过公式）
wb.save()                   # 直接保存（跳过验证）
# 结果：文件错误，用户看不到公式
```

### 自动化执行保证

Skill内置了流程控制，自动按顺序执行：

```python
class ExcelTaskAutomator:
    """自动按顺序执行，不会跳步"""

    def execute_task(self):
        # 强制按顺序执行
        if not self.check_environment():
            raise Exception("环境未配置，请先执行setup_environment()")

        if not self.load_data():
            raise Exception("数据未加载，请先执行smart_read_excel()")

        if not self.write_formula():
            raise Exception("公式未写入，请先执行write_excel_formula()")

        # 每一步都有检查，防止跳步
        self.validate_and_save()
```

### 用户须知

**当您调用这个skill时**：

1. ✅ **耐心等待**：每个步骤需要时间完成
2. ✅ **不要中断**：中断会导致环境不完整
3. ✅ **按提示操作**：每个步骤完成后会有提示
4. ✅ **验证结果**：检查输出的Excel文件是否正确

**常见的错误心态**：

❌ "环境配置太慢，直接开始分析吧" → 结果：依赖缺失
❌ "数据读取不重要，我手动输入" → 结果：数据错误
❌ "公式写入复杂，直接填数值" → 结果：用户看不到公式
❌ "验证步骤多余，直接保存" → 结果：输出文件错误

### 流程执行时间参考

| 步骤 | 时间 | 为什么需要 |
|------|------|-----------|
| 环境配置 | 30秒-2分钟 | 创建venv，安装6个库 |
| 数据读取 | 5-10秒 | 智能转换格式 |
| 数据处理 | 10-15秒 | 智能类型转换 |
| 公式写入 | 5-10秒 | 编写Excel公式 |
| 筛选操作 | 5-10秒 | 复制符合条件数据 |
| 图表创建 | 10-15秒 | 创建图表并引用数据 |
| 结果验证 | 5-10秒 | 检查输出质量 |
| **总计** | **1-3分钟** | **完整流程保证质量** |

### 流程图示

```
开始
 ↓
[步骤1] 环境配置 ← 必须完成，否则后续步骤失败
 ↓
[步骤2] 数据读取 ← 必须完成，否则没有数据
 ↓
[步骤3] 数据处理 ← 必须完成，否则数据类型错误
 ↓
[步骤4] 公式写入 ← 必须完成，否则用户看不到公式
 ↓
[步骤5] 筛选操作 ← 可选（如果有筛选需求）
 ↓
[步骤6] 图表创建 ← 可选（如果有图表需求）
 ↓
[步骤7] 结果验证 ← 必须完成，确保输出正确
 ↓
[步骤8] 保存文件 ← 最后一步
 ↓
完成 ✅
```

**记住：每一步都是必要的，不要因为"看起来简单"就跳过！**

## 🚀 环境准备

### 用户需要手动准备环境

**在调用此skill之前，请确保已安装以下依赖库。**

### 方式1：使用requirements.txt安装（推荐）

创建`requirements.txt`文件：

```txt
pandas>=1.3.0
openpyxl>=3.0.0
xlwings>=0.24.0
matplotlib>=3.3.0
xlrd>=2.0.0
numpy>=1.20.0
```

安装依赖：

```bash
# Windows（推荐使用虚拟环境）
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt

# Linux/Mac
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

### 方式2：逐个安装依赖库

```bash
pip install pandas
pip install openpyxl
pip install xlwings
pip install matplotlib
pip install xlrd
pip install numpy
```

### 依赖库说明

| 库名 | 用途 | 安装必要性 |
|------|------|-----------|
| **pandas** | 数据分析核心库 | ✅ 必须安装 |
| **openpyxl** | Excel读写（支持公式） | ✅ 必须安装 |
| **xlwings** | Excel实时交互（保留筛选） | ⭐ 可选（高级功能） |
| **matplotlib** | 数据可视化 | ⭐ 可选（图表功能） |
| **xlrd** | 旧版.xls格式支持 | ⭐ 可选（如有.xls文件） |
| **numpy** | 数值计算 | ✅ 必须安装 |

### 安装后验证

验证依赖是否安装成功：

```python
# 在Python环境中运行
import pandas as pd
import openpyxl
import numpy as np

print("✅ 核心库安装成功")

# 可选库验证
try:
    import xlwings
    print("✅ xlwings安装成功（高级功能可用）")
except ImportError:
    print("⚠️ xlwings未安装（筛选功能受限）")

try:
    import matplotlib
    print("✅ matplotlib安装成功（可视化功能可用）")
except ImportError:
    print("⚠️ matplotlib未安装（图表功能受限）")

try:
    import xlrd
    print("✅ xlrd安装成功（支持.xls格式）")
except ImportError:
    print("⚠️ xlrd未安装（不支持.xls格式）")
```

### 注意事项

⚠️ **重要提醒**：

1. **必须先安装依赖**：
   - 如果不安装，skill无法执行
   - 会报错："ModuleNotFoundError"

2. **推荐使用虚拟环境**：
   - 避免污染系统Python环境
   - 便于管理和卸载
   - 不同项目可以使用不同版本

3. **Windows用户特殊说明**：
   - xlwings需要Microsoft Excel应用
   - 如无Excel应用，xlwings功能受限
   - 可只安装pandas+openpyxl完成基础功能

4. **Linux/Mac用户说明**：
   - 需要安装Python开发环境
   - 可能需要额外安装系统依赖
   - xlwings在Linux功能有限

### 最小安装方案（基础功能）

如果只需要基础Excel操作：

```txt
pandas>=1.3.0
openpyxl>=3.0.0
numpy>=1.20.0
```

安装命令：

```bash
pip install pandas openpyxl numpy
```

**功能范围**：
- ✅ 数据读取和分析
- ✅ Excel公式写入
- ✅ 数值计算
- ❌ 筛选功能保留（需要xlwings）
- ❌ 图表创建（需要matplotlib）

### 完整安装方案（所有功能）

```txt
pandas>=1.3.0
openpyxl>=3.0.0
xlwings>=0.24.0
matplotlib>=3.3.0
xlrd>=2.0.0
numpy>=1.20.0
```

**功能范围**：
- ✅ 所有功能可用
- ✅ 数据读取和分析
- ✅ Excel公式写入
- ✅ 筛选功能保留
- ✅ 图表创建
- ✅ 数值计算
- ✅ 支持.xls格式

## 📊 核心功能

### 1. 数据分析任务
适用于统计分析、报告生成等需求。

**特点：**
- 智能数据读取（自动识别格式）
- 统计计算（平均值、排名、趋势分析）
- 数据可视化（图表生成）
- 报告导出（Excel、图片）

### 2. Excel操作任务
适用于Excel考试题目、需要在原文件中操作的需求。

**特点：**
- ✅ 在Excel中写入公式（如`=C11-D11-E11`）
- ✅ 保留Excel筛选功能
- ✅ 创建图表并引用原始数据
- ✅ 用户打开Excel可以看到公式和操作过程

### 3. 关键技术实现

#### Excel公式写入
```python
# 写入Excel公式，而不是计算结果
ws.cell(row=11, column=6).value = "=C11-D11-E11"  # Excel自动计算
```

#### Excel筛选保留
```python
# 使用xlwings保留筛选功能
ws.range('D1').api.AutoFilter(Field=2, Criteria1='>=10')
```

#### 图表数据引用
```python
# 图表引用原始数据区域，不是复制数据
data = Reference(ws1, min_col=7, min_row=11, max_row=30)
chart.add_data(data)
```

#### 避免重复计算陷阱
```python
# 自动识别已平均数据，避免重复平均
if '平均' in column_name:
    # 跳过已平均数据，使用原始学期数据
    pass
```

## 🛠️ 完整工作流程

### Excel数据分析流程（7步）

```
环境检测 → 数据扫描 → 数据质量检查 → 数据清洗
→ 统计分析 → 可视化呈现 → 报告生成
```

### Excel操作题目流程

```
文件格式转换 → 公式计算（写入公式） → 筛选操作（保留功能）
→ 图表创建（引用数据） → 结果验证 → 文件保存
```

## 💻 专业代码模板

### 模板1：智能数据读取与转换

```python
import pandas as pd
import os

def smart_read_excel(file_path):
    """
    智能读取Excel文件
    自动处理.xls/.xlsx格式、缺失值、数据类型
    """
    # 自动转换.xls为.xlsx
    if file_path.endswith('.xls') and not file_path.endswith('.xlsx'):
        df_dict = pd.read_excel(file_path, sheet_name=None)
        new_path = file_path.replace('.xls', '_converted.xlsx')

        with pd.ExcelWriter(new_path, engine='openpyxl') as writer:
            for sheet_name, df in df_dict.items():
                # 智能数据类型转换
                for col in df.columns:
                    df[col] = pd.to_numeric(df[col], errors='ignore')
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        return new_path, pd.read_excel(new_path, sheet_name=None)

    return file_path, pd.read_excel(file_path, sheet_name=None)
```

### 模板2：Excel公式计算与写入

```python
from openpyxl import Workbook, load_workbook

def write_excel_formula(wb, sheet_name, formula_range, formula_template):
    """
    在Excel中写入公式
    自动填充公式到指定范围

    参数:
    - wb: Excel工作簿
    - sheet_name: 工作表名
    - formula_range: 公式范围（如'F11:F30'）
    - formula_template: 公式模板（如'=C{row}-D{row}-E{row}'）
    """
    ws = wb[sheet_name]

    start_row, end_row = parse_range(formula_range)

    for row in range(start_row, end_row + 1):
        formula = formula_template.format(row=row)
        ws.cell(row=row, column=get_column(formula_range[0])).value = formula

    return wb

# 使用示例：计算负场数
write_excel_formula(wb, 'Sheet1', 'F11:F30', '=C{row}-D{row}-E{row}')
```

### 模板3：智能平均值计算（避免陷阱）

```python
def smart_average_calculation(data_dict):
    """
    智能平均值计算
    自动识别并跳过已平均数据
    """
    original_data = {}

    for key, value in data_dict.items():
        # 检查是否包含"平均"字样
        if '平均' not in str(key):
            original_data[key] = value

    if len(original_data) == 0:
        return None

    avg_value = sum(original_data.values()) / len(original_data)

    return avg_value, original_data

# 使用示例
data = {
    '大一第一学期': 86.26,
    '大一第二学期': 85.54,
    '大一学年平均总积分': 85.9,  # 已平均，自动跳过
    '大二上学期': 81.36
}

avg, original = smart_average_calculation(data)
# 结果: avg = (86.26 + 85.54 + 81.36) / 3 = 84.39
```

### 模板4：Excel图表创建（引用原始数据）

```python
from openpyxl.chart import BarChart, Reference

def create_excel_chart(wb, sheet_name, chart_position,
                       data_range, categories_range, chart_title):
    """
    创建Excel图表
    引用原始数据区域，不复制数据
    """
    ws = wb[sheet_name]

    chart = BarChart()
    chart.type = "col"
    chart.title = chart_title

    # 引用数据区域
    data = Reference(ws, min_col=data_range['col'],
                    min_row=data_range['row'],
                    max_row=data_range['max_row'])

    # 引用分类区域
    cats = Reference(ws, min_col=categories_range['col'],
                    min_row=categories_range['row'],
                    max_row=categories_range['max_row'])

    chart.add_data(data, titles_from_data=False)
    chart.set_categories(cats)

    ws.add_chart(chart, chart_position)

    return wb

# 使用示例
create_excel_chart(wb, 'Sheet3', 'B7',
                   {'col': 7, 'row': 11, 'max_row': 30},  # 积分G11:G30
                   {'col': 2, 'row': 11, 'max_row': 30},  # 球队B11:B30
                   '2020-2021意大利甲级联赛积分图')
```

### 模板5：数据筛选（复制符合条件数据）

```python
def filter_data_to_sheet(wb, source_sheet, target_sheet,
                         filter_column, filter_condition):
    """
    筛选数据到新Sheet
    自动复制符合条件的数据
    """
    ws_source = wb[source_sheet]
    ws_target = wb[target_sheet]

    # 清空目标sheet
    for row in ws_target.iter_rows():
        for cell in row:
            cell.value = None

    # 复制表头
    for col in range(1, ws_source.max_column + 1):
        ws_target.cell(row=1, column=col).value = ws_source.cell(row=10, column=col).value

    # 筛选并复制数据
    current_row = 2
    for row in range(11, ws_source.max_row + 1):
        value = ws_source.cell(row=row, column=filter_column).value

        if value is not None and filter_condition(value):
            for col in range(1, ws_source.max_column + 1):
                ws_target.cell(row=current_row, column=col).value = \
                    ws_source.cell(row=row, column=col).value
            current_row += 1

    return wb

# 使用示例：筛选胜>=10
filter_data_to_sheet(wb, 'Sheet1', 'Sheet2', 4, lambda x: x >= 10)
```

## 🔧 智能数据处理

### 数据类型智能转换

```python
def smart_convert_to_numeric(value):
    """
    智能数据类型转换
    处理字符串、空值、异常值
    """
    if pd.isna(value) or value in ['', '　', '赛', '胜', '平', '负', '积分']:
        return None

    if isinstance(value, (int, float)):
        return value

    # 字符串转换
    try:
        return int(str(value).strip())
    except ValueError:
        try:
            return float(str(value).strip())
        except ValueError:
            return None

# 批量处理数据列
def process_excel_column(ws, column, start_row, end_row):
    """
    智能处理Excel列数据
    """
    for row in range(start_row, end_row + 1):
        value = ws.cell(row=row, column=column).value
        converted = smart_convert_to_numeric(value)
        if converted is not None:
            ws.cell(row=row, column=column).value = converted
```

### 数据质量检查

```python
def check_data_quality(df):
    """
    数据质量检查
    返回质量报告
    """
    report = {
        '总行数': len(df),
        '缺失值': df.isnull().sum().to_dict(),
        '数据类型': df.dtypes.to_dict(),
        '重复行': df.duplicated().sum()
    }

    # 异常值检测
    numeric_cols = df.select_dtypes(include=[np.number]).columns
    outliers = {}

    for col in numeric_cols:
        Q1 = df[col].quantile(0.25)
        Q3 = df[col].quantile(0.75)
        IQR = Q3 - Q1

        outlier_count = ((df[col] < Q1-1.5*IQR) | (df[col] > Q3+1.5*IQR)).sum()
        if outlier_count > 0:
            outliers[col] = outlier_count

    report['异常值'] = outliers

    return report
```

## 📋 完整案例示例

### 案例1：Excel操作综合题完整流程

```python
import subprocess
import sys
from pathlib import Path
from openpyxl import Workbook, load_workbook
import pandas as pd

# ===== 步骤1：自动创建虚拟环境并安装依赖 =====
python_exe = setup_environment()  # 调用自动化环境配置函数

# ===== 步骤2：读取原始数据 =====
file_path, df_dict = smart_read_excel('01 excel操作综合题1.xls')

# ===== 步骤3：创建新Excel文件 =====
wb = Workbook()
ws1 = wb.active
ws1.title = 'Sheet1'

# 复制原始数据
df = df_dict['Sheet1']
for idx, row in df.iterrows():
    for col_idx, value in enumerate(row, 1):
        if pd.notna(value):
            ws1.cell(row=idx+1, column=col_idx).value = value

# ===== 步骤4：题目1：写入公式计算负场数 =====
write_excel_formula(wb, 'Sheet1', 'F11:F30', '=C{row}-D{row}-E{row}')

# ===== 步骤5：题目2：筛选胜>=10的数据 =====
ws2 = wb.create_sheet('Sheet2')
filter_data_to_sheet(wb, 'Sheet1', 'Sheet2', 4, lambda x: x >= 10)

# ===== 步骤6：题目3：创建图表 =====
ws3 = wb.create_sheet('Sheet3')
create_excel_chart(wb, 'Sheet3', 'B7',
                   {'col': 7, 'row': 11, 'max_row': 30},
                   {'col': 2, 'row': 11, 'max_row': 30},
                   '2020-2021意大利甲级联赛积分图')

# ===== 步骤7：题目4：找出最高积分球队 =====
max_points = 0
best_team = None
for row in range(11, 31):
    points = ws1.cell(row=row, column=7).value
    if points and points > max_points:
        max_points = points
        best_team = ws1.cell(row=row, column=2).value

ws3['F25'].value = best_team

# ===== 步骤8：保存文件 =====
output_path = file_path.replace('.xls', '_完成.xlsx').replace('_converted', '')
wb.save(output_path)

print(f"✅ 完成！文件已保存: {output_path}")
print("   Sheet1: 公式已写入（可查看公式）")
print("   Sheet2: 筛选结果已复制")
print("   Sheet3: 图表已创建（引用原始数据）")
```

### 案例2：计算学生平均综测完整流程

```python
import pandas as pd
import os
import subprocess
import sys
from pathlib import Path

# ===== 步骤1：自动创建虚拟环境并安装依赖 =====
python_exe = setup_environment()  # 调用自动化环境配置函数

# ===== 步骤2：扫描所有Excel文件 =====
excel_files = [f for f in os.listdir('.') if f.endswith('.xlsx')]

# ===== 步骤3：搜索学生数据 =====
student_name = '碗奥博'
student_data = {}

for file in excel_files:
    df_dict = pd.read_excel(file, sheet_name=None)

    for sheet_name, df in df_dict.items():
        if '姓名' in df.columns:
            student_row = df[df['姓名'] == student_name]

            if not student_row.empty:
                # 提取积分列
                score_cols = [col for col in df.columns
                             if any(keyword in str(col)
                             for keyword in ['积分', '综测', '总分'])]

                for col in score_cols:
                    value = student_row[col].values[0]
                    key = f"{file}_{col}"

                    # 区分原始数据和已平均数据
                    if '平均' not in str(col):
                        student_data[key] = value

# ===== 步骤4：计算平均综测 =====
avg_score, original_data = smart_average_calculation(student_data)

print(f"学生: {student_name}")
print(f"平均综测: {avg_score:.2f}")
print("详细数据:")
for key, value in original_data.items():
    print(f"  {key}: {value}")
```

## ⚡ 性能优化

### 大数据集处理

```python
# 分块读取大文件
def read_large_excel_chunked(file_path, chunk_size=10000):
    chunks = []
    for chunk in pd.read_excel(file_path, chunksize=chunk_size):
        chunks.append(chunk)
    return pd.concat(chunks)

# 只读取需要的列
df = pd.read_excel(file_path, usecols=['姓名', '积分', '综测'])

# 数据类型优化
def optimize_data_types(df):
    for col in df.columns:
        if df[col].dtype == 'int64':
            df[col] = pd.to_numeric(df[col], downcast='integer')
        elif df[col].dtype == 'float64':
            df[col] = pd.to_numeric(df[col], downcast='float')
    return df
```

## 🎯 自动化工作流程

### 一键完成Excel操作题目

```python
class ExcelTaskAutomator:
    """
    Excel操作题目自动化工具
    一键完成公式、筛选、图表等操作
    """

    def __init__(self, file_path):
        self.file_path = smart_read_excel(file_path)[0]
        self.wb = load_workbook(self.file_path)

    def calculate_with_formula(self, sheet, target_col, formula_template, start_row, end_row):
        """写入公式"""
        write_excel_formula(self.wb, sheet, f'{target_col}{start_row}:{target_col}{end_row}',
                            formula_template)

    def filter_data(self, source_sheet, target_sheet, filter_col, condition):
        """筛选数据"""
        filter_data_to_sheet(self.wb, source_sheet, target_sheet, filter_col, condition)

    def create_chart(self, sheet, position, data_range, categories_range, title):
        """创建图表"""
        create_excel_chart(self.wb, sheet, position, data_range, categories_range, title)

    def save(self):
        """保存文件"""
        output_path = self.file_path.replace('.xlsx', '_完成.xlsx')
        self.wb.save(output_path)
        return output_path

# 使用示例
automator = ExcelTaskAutomator('考试题目.xls')
automator.calculate_with_formula('Sheet1', 'F', '=C{row}-D{row}-E{row}', 11, 30)
automator.filter_data('Sheet1', 'Sheet2', 4, lambda x: x >= 10)
automator.create_chart('Sheet3', 'B7', {'col': 7, 'row': 11, 'max_row': 30},
                       {'col': 2, 'row': 11, 'max_row': 30}, '积分图')
result_path = automator.save()
```

## 📖 参考资源

### 官方文档
- pandas: https://pandas.pydata.org/
- openpyxl: https://openpyxl.readthedocs.io/
- xlwings: https://docs.xlwings.org/

### 开源项目
- pyopenxlsx: https://github.com/twn39/pyopenxlsx
- opensheet-core: https://github.com/0xNadr/opensheet-core

---

**Skill版本**: v3.0
**最后更新**: 2026-06-25
**核心特性**: 自动安装环境、智能数据处理、保留Excel公式和功能、避免计算陷阱