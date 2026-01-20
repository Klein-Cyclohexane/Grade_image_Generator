# 成绩单图片生成器

从Excel文件读取成绩数据，生成格式化的PNG成绩单图片。

## 功能特点

- 从Excel文件读取成绩数据
- 根据等级自动配色（A+, A/A-, B+/B/B-, C+/C/C-, F, P, NP）
- 自动计算总学分、总均绩、科目数量
- 生成成绩单图片(原始配色参考自北大树洞，可自行修改)

## 依赖

pandas>=2.0.0
Pillow>=10.0.0
openpyxl>=3.1.0


## Excel文件格式要求

Excel文件必须包含以下列（列名必须完全匹配）：

- **科目**: 课程名称
- **代号**: 课程代码
- **等级**: 成绩等级（A+, A, A-, B+, B, B-, C+, C, C-, F, P, NP）
- **绩点**: 对应的GPA绩点（数值）满绩为4.00
- **学分**: 课程学分（数值）
- **学科类型**: 如"专业必修"、"专业选修"、"通选课"等
- **授课教师**: 教师姓名
- **开课院系**: 开课院系名称

## 使用方法

### 基本用法

```bash
python grade_image_generator.py <excel文件路径>
```

### 指定学期

```bash
python grade_image_generator.py grades.xlsx 25Fall
```

### 指定输出文件名称

```bash
python grade_image_generator.py grades.xlsx 25Fall output.png
```


## 输出说明

生成的图片包含：

1. **题头部分**：
   - 学期名称（如"25Fall"）
   - 总学分和计入成绩的科目数量
   - 总均绩（GPA）

2. **课程卡片**：
   - 左侧：学分
   - 中间：课程名称、科目代号、学科类型、授课教师、开课院系
   - 右侧：等级和绩点（P/NP显示"通过"/"不通过"）

## 示例

```bash
python grade_image_generator.py my_grades.xlsx 25Fall grade_report.png
```

## 注意事项

- Excel文件必须是.xlsx格式
- 等级使用标准格式（A+, A, A-, B+, B, B-, C+, C, C-, F, P, NP），无法识别分数成绩（orz）
- P和NP等级的课程不计入GPA计算
- 确保所有数值列（绩点、学分）为数字格式
