"""BOM文件转换模块

此模块提供了将Cadence BOM文件转换为标准格式的功能。
支持HTML和Excel格式的输入文件，输出统一的Excel格式。
"""

from typing import List, Dict, Union, Optional
from pathlib import Path
from bs4 import BeautifulSoup
import pandas as pd

# 定义类型别名
DataFrame = pd.DataFrame

class BOMConverter:
    """BOM文件转换器类
    
    用于将不同格式的BOM文件转换为标准格式。
    支持HTML和Excel格式的输入文件。
    
    Attributes:
        file_path (Path): BOM文件的路径
        file_type (str): 文件类型（'html', 'htm', 'xlsx', 'xls'）
        required_columns (List[str]): 必需的数据列名列表
    """
    
    def __init__(self, file_path: Union[str, Path], file_type: str):
        self.file_path = Path(file_path)
        self.file_type = file_type.lower()
        self.required_columns: List[str] = ['REFDES', 'COMP_VALUE', 'COMP_PACKAGE', 'SYM_MIRROR']

    def convert(self) -> DataFrame:
        try:
            # 读取文件
            if self.file_type in ['htm', 'html']: 
                html_content = self._read_html_file()
                soup = BeautifulSoup(html_content, 'lxml')
                df = self._extract_table_data(soup)
            elif self.file_type in ['xlsx', 'xls']:
                df = self._read_excel_file()
            else:
                raise ValueError(f"Unsupported file type: {self.file_type}")
            
            # 统一列名格式
            df.columns = [col.strip().upper() for col in df.columns]
            
            # 检查所需的列是否存在
            self._check_required_columns(df)
            
            # 过滤所需的列
            df = df[self.required_columns]
            
            # 创建结果DataFrame
            result_df = self._process_data(df)
            
            return result_df

        except Exception as e:
            raise Exception(f"读取或处理文件时发生错误: {e}")

    def _read_html_file(self) -> str:
        with open(self.file_path, 'r', encoding='utf-8') as file:
            return file.read()

    def _extract_table_data(self, soup: BeautifulSoup) -> DataFrame:
        table = soup.find('table')
        if table:
            return pd.read_html(str(table), header=0)[0]
        else:
            raise ValueError("No table found in the HTML file")

    def _read_excel_file(self) -> DataFrame:
        # 读取Excel文件，不指定header
        df = pd.read_excel(self.file_path, header=None)
        
        # 查找有效的表头行
        header_row = -1
        for idx, row in df.iterrows():
            # 将行数据转换为大写进行比较
            row = [str(cell).strip().upper() for cell in row]
            # 检查是否包含所有必需的列
            if all(col in row for col in self.required_columns):
                header_row = idx
                break
        
        if header_row == -1:
            raise ValueError("未找到包含所需列的有效表头行")
        
        # 使用找到的表头行重新读取数据
        df = pd.read_excel(self.file_path, header=header_row)
        return df
    def _check_required_columns(self, df: DataFrame) -> None:
        missing_columns = [col for col in self.required_columns if col not in df.columns]
        if missing_columns:
            raise ValueError(f"以下列未找到: {', '.join(missing_columns)}")

    def _process_data(self, df: DataFrame) -> DataFrame:
        result_df = pd.DataFrame(columns=['COMP_VALUE', 'COMP_PACKAGE', 'Top REFDES', 'Bootom REFDES'])

        for _, row in df.iterrows():
            comp_value = row['COMP_VALUE']
            comp_package = row['COMP_PACKAGE']
            sym_mirror = row['SYM_MIRROR']
            refdes = row['REFDES']

            existing_row = result_df[(result_df['COMP_VALUE'] == comp_value) & (result_df['COMP_PACKAGE'] == comp_package)]

            if existing_row.empty:
                new_row = pd.DataFrame([{
                    'COMP_VALUE': comp_value,
                    'COMP_PACKAGE': comp_package,
                    'Top REFDES': refdes if sym_mirror == 'NO' else '',
                    'Bootom REFDES': refdes if sym_mirror == 'YES' else ''
                }])
                result_df = pd.concat([result_df, new_row], ignore_index=True)
            else:
                index = existing_row.index[0]
                if sym_mirror == 'YES':
                    result_df.at[index, 'Bootom REFDES'] = self._concat_values(result_df.at[index, 'Bootom REFDES'], refdes)
                else:
                    result_df.at[index, 'Top REFDES'] = self._concat_values(result_df.at[index, 'Top REFDES'], refdes)

        result_df['Quantity'] = result_df.apply(lambda row: self._count_non_empty_values(row['Top REFDES'], row['Bootom REFDES']), axis=1)
        result_df.insert(0, 'Item', range(1, len(result_df) + 1))

        return result_df

    def _concat_values(self, current_value: Optional[str], new_value: str) -> str:
        return f"{current_value},{new_value}" if current_value else new_value

    def _count_non_empty_values(self, top_refdes: str, bottom_refdes: str) -> int:
        top_count = len(top_refdes.split(',')) if top_refdes else 0
        bottom_count = len(bottom_refdes.split(',')) if bottom_refdes else 0
        return top_count + bottom_count

def convert_file(file_path: Union[str, Path]) -> DataFrame:
    file_extension = file_path.split('.')[-1].lower()
    supported_extensions = ['xls', 'xlsx', 'html', 'htm']
    
    if file_extension not in supported_extensions:
        raise ValueError(f"Unsupported file extension: {file_extension}")
    
    file_type = file_extension
    converter = BOMConverter(file_path, file_type)
    return converter.convert()