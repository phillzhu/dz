import streamlit as st
import pandas as pd
import re
import os
from io import BytesIO

def extract_order_id(text):
    if pd.isna(text):
        return None
    match = re.search(r'\d{12,}', str(text))
    return match.group(0) if match else None

def process_dataframe(df, source, order_column='订单编号'):
    df['来源'] = source
    if source == '明细':
        df[order_column] = df[order_column].apply(
            lambda x: str(int(float(x))) if pd.notna(x) else None
        )
    else:
        df.insert(0, order_column, df['记录摘要'].apply(extract_order_id))
        df[order_column] = df[order_column].astype(str)
    return df.dropna(subset=[order_column])

def main():
    st.title("京东订单处理工具")
    
    uploaded_file = st.file_uploader("请选择Excel文件", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        try:
            # 读取文件
            df_detail = pd.read_excel(uploaded_file, sheet_name='明细')
            df_done = pd.read_excel(uploaded_file, sheet_name='已做单')
            
            progress_text = st.empty()
            progress_bar = st.progress(0)
            
            # 处理数据
            progress_text.text("正在处理数据...")
            progress_bar.progress(20)
            
            df_detail = process_dataframe(df_detail, '明细')
            df_done = process_dataframe(df_done, '已做单')
            
            progress_bar.progress(40)
            
            # 统计原始数据
            original_detail_count = len(df_detail)
            original_done_count = len(df_done)
            
            # 查找重复订单
            common_order_ids = set(df_detail['订单编号']).intersection(set(df_done['订单编号']))
            
            progress_bar.progress(60)
            
            # 处理重复订单
            df_duplicate = pd.concat([
                pd.concat([df_done[df_done['订单编号'] == order_id],
                          df_detail[df_detail['订单编号'] == order_id]])
                for order_id in common_order_ids
            ], ignore_index=True)
            
            # 处理非重复订单
            non_duplicate_detail = df_detail[~df_detail['订单编号'].isin(common_order_ids)]
            non_duplicate_done = df_done[~df_done['订单编号'].isin(common_order_ids)]
            
            progress_bar.progress(80)
            
            # 创建输出文件
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_duplicate.to_excel(writer, sheet_name='重复订单', index=False)
                non_duplicate_detail.to_excel(writer, sheet_name='明细非重复', index=False)
                non_duplicate_done.to_excel(writer, sheet_name='已做单非重复', index=False)
            
            # 显示统计信息
            st.success("处理完成！")
            st.write("========== 统计信息 ==========")
            st.write(f"原明细总行数: {original_detail_count}")
            st.write(f"原已做单总行数: {original_done_count}")
            st.write(f"重复订单中的已做单行数: {len(df_duplicate[df_duplicate['来源'] == '已做单'])}")
            st.write(f"重复订单中的明细行数: {len(df_duplicate[df_duplicate['来源'] == '明细'])}")
            st.write(f"明细非重复数量: {len(non_duplicate_detail)}")
            st.write(f"已做单非重复数量: {len(non_duplicate_done)}")
            
            # 提供下载按钮
            st.download_button(
                label="下载处理结果",
                data=output.getvalue(),
                file_name="处理结果.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            progress_bar.progress(100)
            
        except Exception as e:
            st.error(f"处理失败：{str(e)}")

if __name__ == '__main__':
    main()