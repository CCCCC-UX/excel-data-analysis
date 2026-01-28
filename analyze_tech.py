# 【核心功能】生成Excel分析报告（热门技术TOP50+薪资统计TOP50，删除技术组合）
import pandas as pd
import numpy as np
from collections import Counter
import ast
import jieba

# 1. Excel文件读取函数（返回数据+读取日志）
def read_excel_file(file_path):
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        read_log = [
            f"数据文件：{file_path}",
            f"数据规模：{df.shape[0]} 行 × {df.shape[1]} 列",
            f"分析字段：technology_label（技术标签）、minimum_monthly_salary（最低月薪）、maximum_monthly_salary（最高月薪）",
            "读取状态：✅ 成功"
        ]
        print("="*50)
        for line in read_log:
            print(line)
        print("="*50)
        return df, read_log
    except FileNotFoundError:
        read_log = [f"读取状态：❌ 失败（未找到文件 {file_path}）"]
        print(read_log[0])
        return None, read_log
    except Exception as e:
        read_log = [f"读取状态：❌ 失败（错误：{str(e)}）"]
        print(read_log[0])
        return None, read_log

# 2. 核心数据分析函数（TOP50调整+删除技术组合）
def analyze_tech_trends(df):
    all_technologies = []
    tech_salary_mapping = {}

    # 第一步：提取原始数据
    for idx, row in df.iterrows():
        if pd.notna(row['technology_label']) and str(row['technology_label']).strip() not in ['', '[]', 'nan']:
            try:
                tech_list = ast.literal_eval(str(row['technology_label']))
                if isinstance(tech_list, list) and len(tech_list) > 0:
                    all_technologies.extend(tech_list)
                    # 提取有效薪资
                    min_sal = row['minimum_monthly_salary']
                    max_sal = row['maximum_monthly_salary']
                    if pd.notna(min_sal) and pd.notna(max_sal) and min_sal > 0 and max_sal > 0:
                        avg_sal = (min_sal + max_sal) / 2
                        for tech in tech_list:
                            tech_salary_mapping[tech] = tech_salary_mapping.get(tech, []) + [avg_sal]
            except:
                continue

    # -------------------------- 模块1：热门技术TOP50（原TOP20→TOP50） --------------------------
    tech_top50_df = pd.DataFrame(columns=['序号', '技术名称', '出现频次', '占比(%)'])
    if all_technologies:
        tech_counter = Counter(all_technologies)
        top50_data = tech_counter.most_common(50)  # 核心修改：20→50
        total_count = sum(tech_counter.values())
        # 构造DataFrame数据
        rows = []
        for i, (tech, count) in enumerate(top50_data, 1):
            proportion = round((count / total_count) * 100, 2)
            rows.append([i, tech, count, proportion])
        tech_top50_df = pd.DataFrame(rows, columns=['序号', '技术名称', '出现频次', '占比(%)'])
        # 添加汇总行
        summary_row = ['-', '汇总', total_count, '100.00']
        tech_top50_df.loc[len(tech_top50_df)] = summary_row

    # -------------------------- 模块2：技术薪资统计TOP50（原TOP15→TOP50） --------------------------
    salary_stats_df = pd.DataFrame(columns=['序号', '技术名称', '样本量(次)', '平均薪资(元)', '薪资中位数(元)', '薪资标准差(元)'])
    if tech_salary_mapping:
        valid_data = []
        for tech, salaries in tech_salary_mapping.items():
            if len(salaries) >= 10:  # 仍保留样本量≥10的筛选条件
                valid_data.append({
                    '技术名称': tech,
                    '样本量(次)': len(salaries),
                    '平均薪资(元)': round(np.mean(salaries), 2),
                    '薪资中位数(元)': round(np.median(salaries), 2),
                    '薪资标准差(元)': round(np.std(salaries), 2)
                })
        # 核心修改：按平均薪资降序排序，取前50（原15）
        valid_data_sorted = sorted(valid_data, key=lambda x: x['平均薪资(元)'], reverse=True)[:50]
        # 构造DataFrame数据
        rows = []
        for i, data in enumerate(valid_data_sorted, 1):
            rows.append([
                i, data['技术名称'], data['样本量(次)'],
                data['平均薪资(元)'], data['薪资中位数(元)'], data['薪资标准差(元)']
            ])
        salary_stats_df = pd.DataFrame(rows, columns=['序号', '技术名称', '样本量(次)', '平均薪资(元)', '薪资中位数(元)', '薪资标准差(元)'])
        # 添加汇总行
        if valid_data_sorted:
            avg_total_salary = round(np.mean([d['平均薪资(元)'] for d in valid_data_sorted]), 2)
            summary_row = ['-', '汇总', f'共{len(valid_data)}个技术', avg_total_salary, '-', '-']
            salary_stats_df.loc[len(salary_stats_df)] = summary_row

    # -------------------------- 模块3：汇总报告（删除技术组合相关统计） --------------------------
    summary_data = {
        '统计项目': [
            '原始数据总行数', '分析字段数', '提取技术标签总数', 
            '不同技术标签数量', '有薪资数据的技术数量（样本≥10）',
            '分析完成时间'  # 核心修改：删除“统计技术组合总数”项
        ],
        '数值': [
            df.shape[0], df.shape[1], sum(Counter(all_technologies).values()) if all_technologies else 0,
            len(Counter(all_technologies)) if all_technologies else 0,
            len([t for t, s in tech_salary_mapping.items() if len(s)>=10]) if tech_salary_mapping else 0,
            pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')
        ]
    }
    summary_df = pd.DataFrame(summary_data)

    # 终端同步输出进度
    print("\n📊 数据分析完成，各模块数据如下：")
    print(f"1. 热门技术TOP50：{len(tech_top50_df)-1} 条数据（含汇总）")  # 修改：20→50
    print(f"2. 技术薪资统计TOP50：{len(salary_stats_df)-1 if not salary_stats_df.empty else 0} 条数据（含汇总）")  # 修改：15→50

    return {
        'tech_top50': tech_top50_df,       # 修改：top20→top50
        'salary_stats': salary_stats_df,
        'summary': summary_df
    }

# 3. 生成多工作表Excel文件（删除技术组合sheet）
def generate_excel_report(analysis_results, read_log, output_file):
    try:
        # 创建ExcelWriter对象（支持多sheet）
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            # 工作表1：热门技术TOP50（修改名称）
            analysis_results['tech_top50'].to_excel(writer, sheet_name='热门技术TOP50', index=False)
            # 工作表2：技术薪资统计TOP50（修改名称）
            analysis_results['salary_stats'].to_excel(writer, sheet_name='技术薪资统计TOP50', index=False)
            # 工作表3：汇总报告
            analysis_results['summary'].to_excel(writer, sheet_name='汇总报告', index=False)
            # 工作表4：数据读取日志
            log_df = pd.DataFrame(read_log, columns=['数据读取日志'])
            log_df.to_excel(writer, sheet_name='读取日志', index=False)
        
        print(f"\n✅ Excel分析报告生成成功！")
        print(f"📁 文件名：{output_file}")
        print(f"📑 包含工作表：热门技术TOP50、技术薪资统计TOP50、汇总报告、读取日志")  # 删除技术组合sheet说明
        print(f"💡 路径：/workspaces/excel-data-analysis/{output_file}（Codespaces当前目录）")
        return True
    except Exception as e:
        print(f"\n❌ 生成Excel失败：{str(e)}")
        print("💡 排查建议：1. 关闭已打开的同名Excel文件 2. 重启Codespaces 3. 检查文件权限")
        return False

# 4. 主程序：一键执行（读取→分析→生成Excel）
if __name__ == "__main__":
    # 配置文件路径
    INPUT_EXCEL = "1-计算机(33351).xlsx"  # 你的原始数据文件
    OUTPUT_EXCEL = "计算机技术趋势分析结果.xlsx"  # 生成的分析报告文件

    print("🚀 开始执行计算机技术趋势分析（TOP50调整+删除技术组合）...")
    # 步骤1：读取原始数据
    df, read_log = read_excel_file(INPUT_EXCEL)
    
    if df is not None:
        # 步骤2：执行核心分析
        print("\n🔍 开始数据分析...")
        analysis_results = analyze_tech_trends(df)
        
        # 步骤3：生成Excel报告
        print("\n📥 开始生成Excel文件...")
        generate_excel_report(analysis_results, read_log, OUTPUT_EXCEL)
        print("\n🎉 所有流程完成！")
    else:
        # 若读取失败，生成仅含日志的Excel
        print("\n📥 生成错误日志Excel...")
        error_summary = pd.DataFrame({
            '统计项目': ['数据读取状态', '错误原因', '建议'],
            '数值': [read_log[0], read_log[0].split('（')[1].strip('）') if '（' in read_log[0] else '-', '检查原始Excel文件路径/完整性']
        })
        generate_excel_report(
            analysis_results={'tech_top50': pd.DataFrame(), 'salary_stats': pd.DataFrame(), 'summary': error_summary},
            read_log=read_log,
            output_file=OUTPUT_EXCEL
        )
        print("\n❌ 分析终止（原始数据读取失败）")
