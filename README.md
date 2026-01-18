# File Manager

文件整理管理工具

## 功能

- 自动扫描源文件夹中的文件
- 根据文件名匹配关键词，自动分类到对应的公司文件夹
- 支持多种文件格式（PDF、JPG、PNG、JPEG）
- 生成详细的处理记录（Excel）
- 可配置的参数（config.json）

## 使用方法

### 1. 配置文件

编辑 `config.json` 设置参数：

```json
{
  "source_path": "源文件夹路径",
  "target_path": "目标文件夹路径",
  "sub_folder_name": "子文件夹名称",
  "allowed_extensions": [".pdf", ".jpg", ".png", ".jpeg"],
  "log_filename_prefix": "日志文件前缀"
}
```

### 2. 运行程序

```bash
python main.py
```

## 系统要求

- Python 3.9+
- pandas
- openpyxl

## 安装依赖

```bash
pip install -r requirements.txt
```

## 输出

- 生成 Excel 日志文件记录所有处理结果
- 文件自动移动到对应的公司文件夹

## 许可

MIT
