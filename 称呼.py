# # 从list.xlsx中获取好友信息,由昵称生成称呼



# 从list.xlsx中读取D列
import pandas as pd
import numpy as np
import re

def get_proper_title(remark):
    """根据备注生成合适的称呼"""
    # 移除括号内容
    remark = re.sub(r'\(.*?\)', '', remark)
    remark = re.sub(r'（.*?）', '', remark)
    
    # 提取姓名（假设中文姓名2-4个字）
    name_match = re.search(r'[\u4e00-\u9fa5]{2,4}(?:[老师|教授|主任|院长])?$', remark)
    if not name_match:
        return remark  # 如果无法提取姓名，返回原始备注
    
    name = name_match.group()
    
    # 特殊身份处理
    if any(title in remark for title in ['老师', '教授', '主任', '院长', '辅导员']):
        return f"{name}老师"
    
    # 提取姓氏
    surname = name[0]
    
    # 根据上下文判断称呼
    if '学校' in remark or '大学' in remark or '学院' in remark:
        if len(name) <= 3:  # 如果是较短的名字，可能是学生
            return f"{name}同学"
    
    # 默认称呼改为"同学"
    return f"{name}同学"  # 移除了性别判断，统一使用"同学"

def is_male_name(surname):
    """简单判断姓名可能的性别（并不准确，仅作参考）"""
    # 这里可以添加更多的女性常见姓氏
    female_surnames = {'王', '李', '张', '刘', '陈', '杨', '黄', '赵', '吴', '周'}
    return surname not in female_surnames

def process_remarks():
    try:
        # 读取Excel文件
        df = pd.read_excel('list.xlsx')
        
        # 添加新列"称呼"
        df['称呼'] = df['备注'].apply(lambda x: get_proper_title(x) if pd.notna(x) else '')
        
        # 保存回Excel
        df.to_excel('list.xlsx', index=False)
        
        # 打印处理结果
        print("备注处理结果：")
        print("-" * 40)
        for _, row in df.iterrows():
            if pd.notna(row['备注']) and row['备注'].strip():
                print(f"原备注: {row['备注']:<20} -> 称呼: {row['称呼']}")
        
        # 打印统计信息
        valid_count = df['称呼'].notna().sum()
        print("-" * 40)
        print(f"总共处理 {valid_count} 个备注")
        print("处理结果已保存到Excel文件中")
        
    except FileNotFoundError:
        print("错误：找不到 list.xlsx 文件")
    except Exception as e:
        print(f"发生错误：{str(e)}")

if __name__ == '__main__':
    process_remarks()


