import shutil
import re
import os
import json
import logging
import pandas as pd
from pathlib import Path
from datetime import datetime

def load_config(config_file="config.json"):
    """加载配置文件"""
    try:
        with open(config_file, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"错误：找不到配置文件 {config_file}")
        raise
    except json.JSONDecodeError:
        print(f"错误：配置文件格式不正确")
        raise

def organize_files_comprehensive(source_roots, target_root, allowed_extensions, log_filename_prefix, search_keyword):
    # 配置日志
    log_filename = f"{log_filename_prefix}_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename, encoding='utf-8')
        ]
    )
    
    target_path = Path(target_root)
    
    logging.info(f"程序启动")
    logging.info(f"源文件夹列表: {source_roots}")
    logging.info(f"目标文件夹: {target_root}")
    logging.info(f"搜索关键字: {search_keyword}")
    logging.info(f"支持的文件格式: {allowed_extensions}")
    
    # 将 allowed_extensions 列表转换为集合并转换为小写
    allowed_extensions = {ext.lower() for ext in allowed_extensions}
    
    # 用于存储执行记录
    execution_records = []

    # 1. 建立公司名称映射表 (支持 2. 或 4、 等前缀)
    logging.info(f"开始扫描目标文件夹: {target_root}")
    company_map = {}
    prefix_pattern = re.compile(r'^\d+[^a-zA-Z\u4e00-\u9fa5]*')

    for folder in target_path.iterdir():
        if folder.is_dir():
            # 第一步：去掉数字前缀
            clean_name = prefix_pattern.sub('', folder.name)
            # 第二步：如果包含"原："或"原:"，只取之前的部分
            if '原：' in clean_name:
                clean_name = clean_name.split('原：')[0]
            elif '原:' in clean_name:
                clean_name = clean_name.split('原:')[0]
            # 第三步：去掉前后空格
            clean_name = clean_name.strip()
            
            if clean_name:
                company_map[clean_name] = folder

    sorted_keys = sorted(company_map.keys(), key=len, reverse=True)
    logging.info(f"找到 {len(company_map)} 个公司文件夹")
    for clean_name in sorted_keys:
        logging.info(f"  - {clean_name} => {company_map[clean_name].name}")

    # 2. 遍历所有源文件夹中的文件
    logging.info(f"开始扫描源文件夹中的文件 (支持格式: {', '.join(allowed_extensions)})")
    print(f"开始扫描文件 (支持格式: {', '.join(allowed_extensions)})...")
    
    # 遍历每个源文件夹
    for source_root in source_roots:
        source_path = Path(source_root)
        
        if not source_path.exists():
            logging.error(f"错误：找不到源文件夹 {source_root}")
            print(f"错误：找不到源文件夹 {source_root}")
            continue
        
        logging.info(f"正在处理源文件夹: {source_root}")
        print(f"\n正在处理源文件夹: {source_root}")
        
        # 使用 iterdir() 遍历所有文件，然后通过后缀过滤
        for file_path in source_path.iterdir():
            # 排除文件夹，且只处理指定后缀的文件
            if file_path.is_dir() or file_path.suffix.lower() not in allowed_extensions:
                continue

        file_name = file_path.name
        matched_key = None
        status = "跳过"
        dest_path_str = "N/A"
        remark = "未匹配到对应的公司文件夹"

        # 匹配逻辑
        for key in sorted_keys:
            if key in file_name:
                matched_key = key
                break
        
        if matched_key:
            target_company_folder = company_map[matched_key]
            logging.info(f"文件 '{file_name}' 匹配到公司: {matched_key}")
            
            # 在目标公司文件夹中查找包含关键字的子文件夹
            matching_subfolder = None
            
            for item in target_company_folder.iterdir():
                if item.is_dir() and search_keyword in item.name:
                    matching_subfolder = item
                    break
            
            # 如果找到了包含关键字的文件夹就使用它，否则直接放在公司文件夹下
            if matching_subfolder:
                final_dest_dir = matching_subfolder
                logging.info(f"  找到子文件夹: {matching_subfolder.name}")
            else:
                final_dest_dir = target_company_folder
                logging.info(f"  未找到包含'{search_keyword}'的子文件夹，使用公司根目录")
            
            dest_path_str = str(final_dest_dir)
            
            try:
                
                # 执行移动操作
                shutil.move(str(file_path), str(final_dest_dir / file_name))
                
                status = "成功"
                remark = f"已移动至: {target_company_folder.name}"
                logging.info(f"  成功移动到: {dest_path_str}")
                print(f"✅ {file_name} -> {target_company_folder.name}")
            except Exception as e:
                status = "失败"
                remark = f"移动出错: {str(e)}"
                logging.error(f"  移动失败: {str(e)}")
                print(f"❌ {file_name} 移动失败")
        else:
            logging.warning(f"文件 '{file_name}' 未匹配到任何公司文件夹")
        
        # 记录到列表
        execution_records.append({
            "文件名": file_name,
            "文件类型": file_path.suffix.upper(),
            "匹配关键词": matched_key if matched_key else "无",
            "执行状态": status,
            "目标路径": dest_path_str,
            "备注": remark,
            "处理时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })

    # 3. 导出 Excel 日志
    if execution_records:
        df = pd.DataFrame(execution_records)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_filename = f"{log_filename_prefix}_{timestamp}.xlsx"
        
        try:
            df.to_excel(excel_filename, index=False)
            success_count = len(df[df['执行状态']=='成功'])
            fail_count = len(df[df['执行状态']!='成功'])
            
            logging.info(f"处理完成 - 成功: {success_count} 个, 跳过/失败: {fail_count} 个")
            logging.info(f"Excel记录已保存至: {excel_filename}")
            
            print(f"\n" + "="*40)
            print(f"处理总结：")
            print(f"成功: {success_count} 个")
            print(f"跳过/失败: {fail_count} 个")
            print(f"详细记录已保存至: {excel_filename}")
            print("="*40)
        except Exception as e:
            logging.error(f"Excel导出失败: {e}")
            print(f"Excel导出失败: {e}")
    else:
        logging.info("未发现符合条件的文件")
        print("未发现符合条件的文件。")
    
    logging.info("程序执行完成")

# --- 从配置文件读取路径和参数 ---
if __name__ == "__main__":
    try:
        config = load_config("config.json")
        
        SOURCES = config["source_paths"]
        TARGET = config["target_path"]
        SEARCH_KEYWORD = config["search_keyword"]
        ALLOWED_EXT = config["allowed_extensions"]
        LOG_PREFIX = config["log_filename_prefix"]
        
        organize_files_comprehensive(SOURCES, TARGET, ALLOWED_EXT, LOG_PREFIX, SEARCH_KEYWORD)
    except Exception as e:
        # 如果在日志配置之前出错，确保写入一个错误日志文件
        error_log = f"error_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        with open(error_log, 'w', encoding='utf-8') as f:
            f.write(f"{datetime.now().strftime('%Y-%m-%d %H:%M:%S')} - ERROR - 程序执行失败: {str(e)}\n")
            import traceback
            f.write(traceback.format_exc())
        print(f"程序执行失败，错误已记录到: {error_log}")
        print(f"错误信息: {str(e)}")
        raise