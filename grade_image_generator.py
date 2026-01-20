#!/usr/bin/env python
"""
成绩单图片生成器
从Excel文件读取成绩数据，生成格式化的PNG图片
"""

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
import os
import sys
from typing import Dict, Tuple, List
import math


class GradeImageGenerator:
    """成绩单图片生成器类"""
    
    #（RGB）
    # 实色部分（高饱和度）
    GRADE_COLORS_SOLID = {
        'A+': (109,255,77),     
        'A': (169,255,72),      
        'A-': (169,255,72),     
        'B+': (223,255,100),    
        'B': (223,255,100),     
        'B-': (223,255,100),     
        'C+': (255,234,109),     
        'C': (255,234,109),      
        'C-': (255,234,109),     
        'F': (79, 148, 205),      
        'P': (217,217,243),     
        'NP': (79, 148, 205),     
    }
    
    # 非实色部分（低饱和度背景）
    GRADE_COLORS_BACKGROUND = {
        'A+': (118, 222, 96),     # #76de60
        'A': (158,220,95),      # #76de60
        'A-': (158,220,95),     # #76de60
        'B+': (193,221,94),     # #9edd60
        'B': (193,221,94),      # #9edd60
        'B-': (193,221,94),     # #9edd60
        'C+': (222,203,97),     # #deca60
        'C': (222,203,97),      # #deca60
        'C-': (222,203,97),     # #deca60
        'F': (79, 148, 205),      # #4F94CD
        'P': (217,217,243),     # #d9d9f5
        'NP': (79, 148, 205),     # #4F94CD
    }
    
    # 题头颜色
    HEADER_COLOR = (206,255,65)  # #ceff41
    
    # 等级到实心条占比的映射（0-1之间）
    GRADE_FILL_RATIO = {
        'A+': 1.0,    # 100%
        'A': 0.9,     # 90%
        'A-': 0.8,    # 80%
        'B+': 0.7,    # 70%
        'B': 0.6,     # 60%
        'B-': 0.5,    # 50%
        'C+': 0.4,    # 50%
        'C': 0.3,     # 40%
        'C-': 0.2,    # 30%
        'F': 0.0,     # 0%
        'P': 0.0,     # 0%
        'NP': 0.0,    # 0%
    }
    
    # 默认颜色（如果等级不在上述dic中）
    DEFAULT_COLOR = (158, 158, 158)  # 灰色
    
    def reduce_saturation(self, color: Tuple[int, int, int], factor: float = 0.3) -> Tuple[int, int, int]:
        r, g, b = color
        # 计算灰度值
        gray = int(0.299 * r + 0.587 * g + 0.114 * b)
        # 混合原色和灰度
        new_r = int(r * factor + gray * (1 - factor))
        new_g = int(g * factor + gray * (1 - factor))
        new_b = int(b * factor + gray * (1 - factor))
        return (new_r, new_g, new_b)
    
    def __init__(self, excel_path: str, semester_name: str = "25Fall"):
        """
        初始化生成器
        
        Args:
            excel_path: Excel文件路径
            semester_name: 学期名称，如"25Fall"
        """
        self.excel_path = excel_path
        self.semester_name = semester_name
        self.df = None
        self.load_data()
    
    def load_data(self):
        """加载Excel数据"""
        try:
            self.df = pd.read_excel(self.excel_path)
            # 检查列
            required_columns = ['科目', '代号', '等级', '绩点', '学分', 
                              '类型', '授课教师', '开课院系']
            missing_columns = [col for col in required_columns if col not in self.df.columns]
            if missing_columns:
                raise ValueError(f"Excel文件缺少必需的列: {', '.join(missing_columns)}")
        except Exception as e:
            raise Exception(f"读取Excel文件失败: {str(e)}")
    
    def get_grade_color_solid(self, grade: str) -> Tuple[int, int, int]:
        """
        根据等级获取实色部分颜色（高饱和度）
        
        Args:
            grade: 等级字符串（如"A+", "A", "B+"等）
        
        Returns:
            RGB颜色元组
        """
        grade = str(grade).strip().upper()
        
        if grade in self.GRADE_COLORS_SOLID:
            return self.GRADE_COLORS_SOLID[grade]
        
        # 如果不在dic中，返回默认颜色
        return self.DEFAULT_COLOR
    
    def get_grade_color_background(self, grade: str) -> Tuple[int, int, int]:
        """
        根据等级获取背景颜色（低饱和度）
        
        Args:
            grade: 等级字符串（如"A+", "A", "B+"等）
        
        Returns:
            RGB颜色元组
        """
        grade = str(grade).strip().upper()
        
        if grade in self.GRADE_COLORS_BACKGROUND:
            return self.GRADE_COLORS_BACKGROUND[grade]
        
        # 如果不在dic中，返回默认颜色
        return self.DEFAULT_COLOR
    
    def calculate_statistics(self) -> Dict:
        """
        计算统计信息
        
        Returns:
            包含总学分、总均绩、科目数量
        """
        # 过滤掉P和NP的课程（通常不计入GPA），F档计入
        gpa_courses = self.df[~self.df['等级'].isin(['P', 'NP', 'p', 'np'])]
        
        # 计算总学分
        total_credits = self.df['学分'].sum()
        
        # 计算加权平均绩点（F档计入总学分）
        if len(gpa_courses) > 0:
            weighted_gpa = (gpa_courses['绩点'] * gpa_courses['学分']).sum() / gpa_courses['学分'].sum()
            total_gpa = round(weighted_gpa, 2)
        else:
            total_gpa = 0.0
        
        # 计入成绩的科目数量（排除P和NP，但包括F）
        counted_courses = len(gpa_courses)

        
        return {
            'total_credits': total_credits,
            'total_gpa': total_gpa,
            'counted_courses': counted_courses,
            'total_courses': len(self.df)
        }
    
    def get_font(self, size: int, bold: bool = False):
        """
        获取字体（尝试使用系统字体，如果失败则使用默认字体）
        
        Args:
            size: 字体大小
            bold: 是否加粗
        
        Returns:
            ImageFont对象
        """
        try:
            if sys.platform == 'win32':
                if bold:
                    font_path = "C:/Windows/Fonts/msyhbd.ttc"  
                    if not os.path.exists(font_path):
                        font_path = "C:/Windows/Fonts/simhei.ttf"  
                else:
                    font_path = "C:/Windows/Fonts/msyh.ttc"  
                    if not os.path.exists(font_path):
                        font_path = "C:/Windows/Fonts/simsun.ttc"  
                if os.path.exists(font_path):
                    return ImageFont.truetype(font_path, size)
        except:
            pass
        
        # 如果都失败，使用默认字体
        try:
            return ImageFont.truetype("arial.ttf", size)
        except:
            return ImageFont.load_default()
    
    def _get_grade_sort_order(self, grade: str) -> int:
        """
        取等级排序顺序
        
        Args:
            grade: 等级字符串
        
        Returns:
            排序顺序数字
        """
        grade = str(grade).strip().upper()
        order_map = {
            'A+': 1,
            'A': 2,
            'A-': 3,
            'B+': 4,
            'B': 5,
            'B-': 6,
            'C+': 7,
            'C': 8,
            'C-': 9,
            'P': 10,
            'F': 11,
            'NP': 12,
        }
        return order_map.get(grade, 99) 
    
    def _sort_grades(self):
        """
        对成绩数据按等级从高到低排序，同等级按学分从高到低排序
        """
        # 创建排序列
        self.df['_sort_order'] = self.df['等级'].apply(self._get_grade_sort_order)
        # 按排序顺序排序：等级升序（高在前），学分降序（高在前）
        self.df = self.df.sort_values(['_sort_order', '学分'], ascending=[True, False])
        # 删除临时排序列
        self.df = self.df.drop('_sort_order', axis=1)
        # 重置索引
        self.df = self.df.reset_index(drop=True)
    
    def generate_image(self, output_path: str = "grade_report.png"):
        """
        生成成绩单图片
        
        Args:
            output_path: 输出图片路径
        """
        # 对成绩按等级排序
        self._sort_grades()
        
        # 计算统计信息
        stats = self.calculate_statistics()
        
        card_width = 800
        card_height = 130  
        card_margin = 15
        header_height = 140
        padding = 20
        
        # 计算总高度
        total_height = header_height + padding + len(self.df) * (card_height + card_margin) + padding
        
        # 创建图片
        img = Image.new('RGB', (card_width, total_height), color=(245, 245, 245))
        draw = ImageDraw.Draw(img)
        
        # 题头
        self._draw_header(draw, stats, card_width, header_height)
        
        # 绘制课程卡片
        y_offset = header_height + padding
        for idx, row in self.df.iterrows():   #idx不需要
            self._draw_course_card(draw, row, 0, y_offset, card_width, card_height)
            y_offset += card_height + card_margin
        
        # 保存图片
        img.save(output_path, 'PNG', quality=95)
        print(f"成绩单图片已生成: {output_path}")
    
    def _draw_header(self, draw: ImageDraw.Draw, stats: Dict, width: int, height: int):
        """
        绘制题头
        
        Args:
            draw: ImageDraw对象
            stats: 统计信息字典
            width: 题头宽度
            height: 题头高度
        """
        # 题头背景色 #ceff41
        draw.rectangle([0, 0, width, height], fill=self.HEADER_COLOR)
        
        # 字体
        font_large = self.get_font(32, bold=True)
        font_medium = self.get_font(24, bold=True)
        font_small = self.get_font(18)
        
        # 学期名称
        semester_text = f"{self.semester_name}"
        draw.text((20, 20), semester_text, fill=(0, 0, 0), font=font_medium)
        
        # 总学分和课程数（左侧显示）
        # 总学分：所有科目的学分总和
        total_credits = stats['total_credits']
        credits_text = f"学分 {total_credits} 共{stats['total_courses']}门课程"
        draw.text((20, 60), credits_text, fill=(0, 0, 0), font=font_small)
        
        # 总均绩（右侧大号显示）
        gpa_text = f"{stats['total_gpa']}"
        bbox = draw.textbbox((0, 0), gpa_text, font=font_large)
        text_width = bbox[2] - bbox[0]
        draw.text((width - text_width - 20, 20), gpa_text, fill=(0, 0, 0), font=font_large)
        
        # 百分制成绩（总均绩下方，满绩4.00换算成百分制）
        # 绩点/4.0 * 100 = 百分制
        percentage_score = stats['total_gpa'] / 4.0 * 100
        percentage_text = f"{percentage_score:.1f}"
        bbox = draw.textbbox((0, 0), percentage_text, font=font_small)
        text_width = bbox[2] - bbox[0]
        draw.text((width - text_width - 20, 70), percentage_text, fill=(0, 0, 0), font=font_small)
    
    def _draw_course_card(self, draw: ImageDraw.Draw, row: pd.Series, 
                         x: int, y: int, width: int, height: int):
        """
        绘制单个课程卡片
        
        Args:
            draw: ImageDraw对象
            row: 课程数据行
            x: 起始x坐标
            y: 起始y坐标
            width: 宽度
            height: 高度
        """
        # 获取等级颜色
        grade = str(row['等级']).strip().upper()
        solid_color = self.get_grade_color_solid(grade)  # 实色
        bg_color = self.get_grade_color_background(grade)  # 背景
        gpa = row['绩点'] if pd.notna(row['绩点']) else 0.0
        
        # 获取该等级对应的实心条占比
        fill_ratio = self.GRADE_FILL_RATIO.get(grade, 0.0)
        
        if fill_ratio >= 1.0:
            # A+类（100%）：全实色
            draw.rectangle([x, y, x + width, y + height], fill=solid_color)
        else:
            # 其他等级：低饱和度背景 + 高饱和度实心条
            # 绘制低饱和度背景
            draw.rectangle([x, y, x + width, y + height], fill=bg_color)
            
            fill_width = int(width * fill_ratio)
            
            # 绘制高饱和度实心条（从左到右）
            if fill_width > 0:
                draw.rectangle([x, y, x + fill_width, y + height], fill=solid_color)
        
        # 字体
        font_large = self.get_font(28, bold=True)
        font_medium = self.get_font(18)
        font_small = self.get_font(14)
        
        text_color = (0, 0, 0)  # 黑色
        
        # 左侧：学分标签和学分数字
        # 先绘制"学分"文字（小号）
        credits_label = "学分"
        draw.text((x + 20, y + 15), credits_label, fill=text_color, font=font_small)
        # 再绘制学分数字（大号）
        credits = str(row['学分'])
        draw.text((x + 20, y + 40), credits, fill=text_color, font=font_large)
        
        # 中间：课程信息
        course_name = str(row['科目']) if pd.notna(row['科目']) else "未知课程"
        course_code = str(row['代号']) if pd.notna(row['代号']) else ""
        course_type = str(row['类型']) if pd.notna(row['类型']) else ""
        teacher = str(row['授课教师']) if pd.notna(row['授课教师']) else ""
        department = str(row['开课院系']) if pd.notna(row['开课院系']) else ""
        
        # 课程名称
        draw.text((x + 100, y + 15), course_name, fill=text_color, font=font_medium)
        
        # 详细信息
        detail_text = f"{course_type} {teacher} ({department})"
        if len(detail_text) > 50:
            detail_text = detail_text[:47] + "..."
        draw.text((x + 100, y + 50), detail_text, fill=text_color, font=font_small)
        
        # 科目代号（显示在详细信息下方）
        if course_code and str(course_code).lower() != 'nan':
            code_text = str(course_code)
            if len(code_text) > 30:
                code_text = code_text[:27] + "..."
            draw.text((x + 100, y + 75), code_text, fill=text_color, font=font_small)
        
        # 右侧：等级和绩点
        grade_text = str(row['等级']) if pd.notna(row['等级']) else ""
        
        # 如果是P或NP，显示"通过"或"不通过"
        if grade in ['P', 'NP']:
            status_text = "通过" if grade == 'P' else "不通过"
            bbox = draw.textbbox((0, 0), grade_text, font=font_large)
            text_width = bbox[2] - bbox[0]
            draw.text((x + width - text_width - 20, y + 20), grade_text, 
                     fill=text_color, font=font_large)
            bbox = draw.textbbox((0, 0), status_text, font=font_small)
            text_width = bbox[2] - bbox[0]
            draw.text((x + width - text_width - 20, y + 60), status_text, 
                     fill=text_color, font=font_small)
        else:
            # 显示等级
            bbox = draw.textbbox((0, 0), grade_text, font=font_large)
            text_width = bbox[2] - bbox[0]
            draw.text((x + width - text_width - 20, y + 20), grade_text, 
                     fill=text_color, font=font_large)
            # 显示绩点
            gpa_text = f"{gpa:.2f}"
            bbox = draw.textbbox((0, 0), gpa_text, font=font_small)
            text_width = bbox[2] - bbox[0]
            draw.text((x + width - text_width - 20, y + 60), gpa_text, 
                     fill=text_color, font=font_small)


def main():
    if len(sys.argv) < 2:
        print("使用方法: python grade_image_generator.py <excel文件路径> [学期名称] [输出文件路径]")
        print("示例: python grade_image_generator.py grades.xlsx 25Fall output.png")
        sys.exit(1)
    
    excel_path = sys.argv[1]
    semester_name = sys.argv[2] if len(sys.argv) > 2 else "25Fall"
    output_path = sys.argv[3] if len(sys.argv) > 3 else "grade_report.png"
    
    if not os.path.exists(excel_path):
        print(f"错误: Excel文件不存在: {excel_path}")
        sys.exit(1)
    
    try:
        generator = GradeImageGenerator(excel_path, semester_name)
        generator.generate_image(output_path)
    except Exception as e:
        print(f"错误: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
