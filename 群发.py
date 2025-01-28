# # 从list.xlsx中获取好友备注和称呼群发祝福
import pandas as pd
import time
import PyOfficeRobot


def send_messages():
    try:
        # 读取Excel文件
        df = pd.read_excel('list.xlsx')
        
        # 确保必要的列存在
        if '备注' not in df.columns or '称呼' not in df.columns:
            print("错误：Excel文件中缺少必要的列（备注/称呼）")
            return
            
        # 遍历每一行数据
        for index, row in df.iterrows():
            if pd.notna(row['备注']) and pd.notna(row['称呼']):
                remark = str(row['备注']).strip()
                title = str(row['称呼']).strip()
                if remark and title:  # 确保都不是空值
                    try:
                        if "同学" in title:
                            message = f"{title}你好，新的一年到来，xxx在此送上诚挚的祝福：愿你在新的一年里，学业进步，生活幸福，心想事成！希望每一天都充满阳光和笑声，无论遇到什么挑战，都能勇敢面对，轻松克服。新的一年，愿你不忘初心，继续前行，成就更多精彩！新年快乐！🎉🎆"
                        elif "老师" in title or "院长" in remark:
                            message = f"{title}您好，学生值此新春佳节之际，衷心祝愿您在新的一年里，身体健康，工作顺利，家庭幸福，万事如意,阖家安康！"
                        elif "总" in remark or "处长" in remark:
                            message = f"{title}您好，新的一年到来，xxx在此送上诚挚的祝福：衷心祝愿您在新的一年里，身体健康，工作顺利，家庭幸福，万事如意,阖家安康！"
                        else:
                            message = f"{title}您好，新的一年到来，xxx老弟在此送上诚挚的祝福：衷心祝愿您在新的一年里，身体健康，工作顺利，家庭幸福，万事如意,阖家安康！"
                        # 这里是发送消息的逻辑
                        print(f"发送给 {remark} 的消息是: {message}")
                        PyOfficeRobot.chat.send_message(who=remark,message=message)
                    except Exception as e:
                        print(f"发送给 {remark} 失败: {str(e)}")
        
        print("消息发送完成！")
        
    except FileNotFoundError:
        print("错误：找不到 list.xlsx 文件")
    except Exception as e:
        print(f"发生错误：{str(e)}")

if __name__ == '__main__':
    send_messages()
