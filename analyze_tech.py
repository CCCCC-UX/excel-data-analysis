# 1. 导入所有必需依赖库（一次性导入，无需额外安装其他库）
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from collections import Counter
import ast
from wordcloud import WordCloud
import jieba

# 2. 读取你的Excel文件（已指定你的文件名，无需修改路径）
def read_excel_file(file_path):
    try:
        # 读取Excel文件（engine='openpyxl'适配.xlsx格式，你的文件格式匹配）
        df = pd.read_excel(file_path, engine='openpyxl')
        print(f"✅ 成功读取Excel文件：{file_path}")
        print(f"📊 数据规模：共 {df.shape[0]} 行，{df.shape[1]} 列")
        # 打印列名（确认与代码中使用的列名一致，你截图显示列名为英文，无需修改）
        print("\n📋 Excel文件列名清单：")
        print(df.columns.tolist())
        return df
    except Exception as e:
        print(f"❌ 读取文件失败：{str(e)}")
        print("💡 请检查：1. 文件是否在当前目录 2. 文件名是否正确")
        return None

# 3. 核心分析函数（技术热度、薪资、组合、词云，全修正）
def analyze_tech_trends(df):
    """分析计算机领域技术趋势：热度TOP20、薪资关联、热门组合、词云"""
    # 3.1 提取技术标签与薪资映射
    all_technologies = []  # 存储所有技术标签
    tech_salary_mapping = {}  # 存储技术与对应薪资的映射
    
    # 遍历每行数据（使用你Excel的英文列名，无需修改）
    for idx, row in df.iterrows():
        # 处理技术标签列（跳过空值或无效值）
        if pd.notna(row['technology_label']) and row['technology_label'] not in ['', '[]']:
            try:
                # 解析字符串格式的列表（如"['Python','Java']"转成实际列表）
                tech_list = ast.literal_eval(row['technology_label']) if isinstance(row['technology_label'], str) else row['technology_label']
                if isinstance(tech_list, list):
                    all_technologies.extend(tech_list)  # 加入所有技术标签
                    
                    # 计算平均薪资（最低+最高月薪取平均）
                    avg_salary = (row['minimum_monthly_salary'] + row['maximum_monthly_salary']) / 2
                    # 为每个技术标签记录薪资
                    for tech in tech_list:
                        if tech not in tech_salary_mapping:
                            tech_salary_mapping[tech] = []
                        tech_salary_mapping[tech].append(avg_salary)
            except:
                continue  # 跳过解析失败的行，不影响整体分析
    
    # 3.2 技术热度排名（TOP20）
    if not all_technologies:
        print("⚠️ 未提取到有效技术标签，无法进行热度分析")
        top_20_tech = []
    else:
        tech_counter = Counter(all_technologies)
        top_20_tech = tech_counter.most_common(20)  # 取出现次数最多的20个技术
        print(f"\n🏆 热门技术技能需求TOP20（按出现频次排序）：")
        for i, (tech, count) in enumerate(top_20_tech, 1):
            print(f"{i:2d}. {tech:<20} 出现 {count:4d} 次")
    
    # 3.3 技术薪资分析（仅统计出现10次以上的技术，避免样本过小）
    tech_salary_stats = {}
    if tech_salary_mapping:
        for tech, salaries in tech_salary_mapping.items():
            if len(salaries) > 10:  # 只保留样本量足够的技术
                tech_salary_stats[tech] = {
                    '出现次数': len(salaries),
                    '平均薪资(元)': round(np.mean(salaries), 2),
                    '薪资中位数(元)': round(np.median(salaries), 2)
                }
        print(f"\n💰 技术薪资统计（样本量≥10的技术）：共 {len(tech_salary_stats)} 个技术")

    # 3.4 热门技术组合分析（TOP15，技术对组合）
    tech_combinations = Counter()
    for idx, row in df.iterrows():
        if pd.notna(row['technology_label']) and row['technology_label'] not in ['', '[]']:
            try:
                tech_list = ast.literal_eval(row['technology_label']) if isinstance(row['technology_label'], str) else row['technology_label']
                if isinstance(tech_list, list) and len(tech_list) >= 2:
                    # 生成有序技术对（避免Python+Java和Java+Python被视为不同组合）
                    for i in range(len(tech_list)):
                        for j in range(i+1, len(tech_list)):
                            combo = tuple(sorted([tech_list[i], tech_list[j]]))
                            tech_combinations[combo] += 1
            except:
                continue
    top_combinations = tech_combinations.most_common(15)
    if top_combinations:
        print(f"\n🔗 热门技术组合TOP15（按出现频次排序）：")
        for i, (combo, count) in enumerate(top_combinations, 1):
            print(f"{i:2d}. {combo[0]} + {combo[1]:<15} 出现 {count:4d} 次")

    # 3.5 可视化模块（全修正：Linux字体+无缩进错误+子图布局）
    plt.rcParams['font.sans-serif'] = ['DejaVu Sans']  # 适配Linux系统字体，避免乱码
    plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示异常
    plt.figure(figsize=(16, 12))  # 整体图表尺寸（宽16，高12，避免子图拥挤）

    # 子图1：热门技术TOP20（横向柱状图，便于查看长技术名称）
    plt.subplot(2, 2, 1)
    if top_20_tech:
        tech_names, tech_counts = zip(*top_20_tech)
        plt.barh(range(len(tech_names)), tech_counts, color='#1f77b4', alpha=0.8)
        plt.yticks(range(len(tech_names)), tech_names, fontsize=9)
        plt.xlabel('出现频次', fontsize=10)
        plt.title('热门技术技能需求TOP20', fontsize=12, fontweight='bold')
        # 添加数值标签（在柱状图右侧显示具体频次）
        for i, count in enumerate(tech_counts):
            plt.text(count + 10, i, str(count), va='center', fontsize=8)
    else:
        plt.text(0.5, 0.5, '无有效技术数据', ha='center', va='center', transform=plt.gca().transAxes)
        plt.title('热门技术技能需求TOP20', fontsize=12, fontweight='bold')

    # 子图2：技术热度vs薪资水平（散点图，关联频次与薪资）
    plt.subplot(2, 2, 2)
    if tech_salary_stats:
        tech_names_list = list(tech_salary_stats.keys())
        tech_freq = [tech_salary_stats[tech]['出现次数'] for tech in tech_names_list]
        tech_salary = [tech_salary_stats[tech]['平均薪资(元)'] for tech in tech_names_list]
        # 绘制散点图（点的大小代表出现次数，颜色区分薪资区间）
        scatter = plt.scatter(tech_freq, tech_salary, c=tech_salary, cmap='YlOrRd', 
                             alpha=0.7, s=[f*0.3 for f in tech_freq])
        # 标注高价值技术（出现>100次 或 薪资>20000元）
        for i, tech in enumerate(tech_names_list):
            if tech_freq[i] > 100 or tech_salary[i] > 20000:
                plt.annotate(tech, (tech_freq[i], tech_salary[i]), 
                            fontsize=8, ha='right', xytext=(5, 0), textcoords='offset points')
        plt.xlabel('出现频次', fontsize=10)
        plt.ylabel('平均薪资（元）', fontsize=10)
        plt.title('技术热度vs薪资水平', fontsize=12, fontweight='bold')
        plt.colorbar(scatter, label='平均薪资（元）')  # 添加颜色条，解释薪资区间
    else:
        plt.text(0.5, 0.5, '无足够薪资数据', ha='center', va='center', transform=plt.gca().transAxes)
        plt.title('技术热度vs薪资水平', fontsize=12, fontweight='bold')

    # 子图3：热门技术组合TOP15（横向柱状图）
    plt.subplot(2, 2, 3)
    if top_combinations:
        combo_names = [f"{combo[0]}\n+{combo[1]}" for combo, count in top_combinations]  # 换行显示长组合名
        combo_counts = [count for combo, count in top_combinations]
        plt.barh(range(len(combo_names)), combo_counts, color='#2ca02c', alpha=0.8)
        plt.yticks(range(len(combo_names)), combo_names, fontsize=8)
        plt.xlabel('组合出现频次', fontsize=10)
        plt.title('热门技术组合TOP15', fontsize=12, fontweight='bold')
        # 添加数值标签
        for i, count in enumerate(combo_counts):
            plt.text(count + 5, i, str(count), va='center', fontsize=8)
    else:
        plt.text(0.5, 0.5, '无有效技术组合数据', ha='center', va='center', transform=plt.gca().transAxes)
        plt.title('热门技术组合TOP15', fontsize=12, fontweight='bold')

    # 子图4：技术关键词词云（适配Linux字体，无需额外安装）
    plt.subplot(2, 2, 4)
    if all_technologies:
        tech_text = ' '.join(all_technologies)  # 拼接所有技术标签为文本
        # 使用Linux系统自带的DejaVuSans字体（路径固定，无需修改）
        wordcloud = WordCloud(
            font_path='/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf',
            width=800, height=400,
            background_color='white',
            max_words=200,  # 最多显示200个关键词
            collocations=False,  # 避免重复显示组合词（如“Python+Java”不重复）
            contour_width=1, contour_color='lightgray'  # 增加边框，更美观
        ).generate(tech_text)
        plt.imshow(wordcloud, interpolation='bilinear')  # bilinear让词云边缘更平滑
        plt.axis('off')  # 隐藏坐标轴，聚焦词云
        plt.title('技术关键词词云', fontsize=12, fontweight='bold')
    else:
        plt.text(0.5, 0.5, '无有效技术标签生成词云', ha='center', va='center', transform=plt.gca().transAxes)
        plt.axis('off')
        plt.title('技术关键词词云', fontsize=12, fontweight='bold')

    # 调整子图间距，避免标题/标签重叠
    plt.tight_layout(pad=3.0)  # pad增加整体边距
    # 保存图表（高清300dpi，避免标签被截断）
    plt.savefig('tech_trends_analysis.png', dpi=300, bbox_inches='tight', facecolor='white')
    plt.show()
    print(f"\n📊 分析图表已保存至当前目录：tech_trends_analysis.png")

    # 返回分析结果，便于后续二次处理（可选）
    return {
        'top_20_technologies': top_20_tech,
        'tech_salary_statistics': tech_salary_stats,
        'top_15_combinations': top_combinations
    }

# 4. 主程序入口（执行读取+分析，一键运行）
if __name__ == "__main__":
    # 你的Excel文件名（固定为你的文件，无需修改）
    EXCEL_FILE = "1-计算机(33351).xlsx"
    # 第一步：读取Excel文件
    df = read_excel_file(EXCEL_FILE)
    # 第二步：若读取成功，执行分析
    if df is not None:
        print("\n🚀 开始执行技术趋势分析...")
        results = analyze_tech_trends(df)
        print("\n✅ 技术趋势分析全部完成！")
        print("📁 生成文件清单：")
        print("1. 分析脚本：analyze_tech.py")
        print("2. 分析图表：tech_trends_analysis.png")
        print("3. 原始数据：1-计算机(33351).xlsx")
    else:
        print("\n❌ 文件读取失败，无法执行分析，请检查文件路径和名称")