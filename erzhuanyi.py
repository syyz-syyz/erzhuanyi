import streamlit as st
import pandas as pd
import io

def convert_2d_to_1d(df):
    """
    将二维 DataFrame 转换为一维格式
    第一列保持不变，后面的列转换为字段名和对应值
    调整逻辑：先遍历完一个字段名称下的所有值，再处理下一个字段名称
    """
    if df.empty:
        return pd.DataFrame()
    
    # 获取第一列的列名
    first_column = df.columns[0]
    
    # 创建一个空的 DataFrame 用于存储结果
    result_df = pd.DataFrame(columns=[first_column, '字段名称', '值内容'])
    
    # 计算总工作量（总行数 * 总列数）
    total_work = (len(df.columns) - 1) * len(df)
    current_work = 0
    
    # 创建进度条
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # 遍历每一个字段（从第二列开始）
    for col_idx in range(1, len(df.columns)):
        field_name = df.columns[col_idx]  # 字段名称为列标题
        
        # 遍历每一行（获取每个字段对应的值）
        for _, row in df.iterrows():
            # 获取第一列的值
            first_value = row[first_column]
            # 获取当前字段的值
            value = row[col_idx]
            
            # 添加到结果 DataFrame
            result_df = pd.concat([result_df, pd.DataFrame({
                first_column: [first_value],
                '字段名称': [field_name],
                '值内容': [value]
            })], ignore_index=True)
            
            # 更新进度
            current_work += 1
            progress_percent = current_work / total_work
            progress_bar.progress(progress_percent)
            status_text.text(f"正在处理: {int(progress_percent * 100)}%")
    
    # 完成后隐藏进度条和状态文本
    progress_bar.empty()
    status_text.empty()
    
    return result_df

def main():
    st.title("Excel 二维转一维转换工具")
    st.markdown("上传一个 Excel 文件，将其从二维格式转换为一维格式。")
    
    # 上传文件
    uploaded_file = st.file_uploader("选择 Excel 文件", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            # 读取 Excel 文件
            df = pd.read_excel(uploaded_file)
            
            # 显示原始数据预览
            st.subheader("原始数据预览")
            st.dataframe(df.head())
            
            # 执行转换
            st.subheader("正在转换数据...")
            converted_df = convert_2d_to_1d(df)
            
            # 显示转换后的数据预览
            st.subheader("转换后的数据预览")
            st.dataframe(converted_df.head())
            
            # 下载转换后的数据
            if not converted_df.empty:
                st.subheader("下载转换后的数据")
                
                # 创建 Excel 文件的二进制流
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    converted_df.to_excel(writer, sheet_name='转换结果', index=False)
                output.seek(0)
                
                # 创建下载按钮
                st.download_button(
                    label="下载 Excel 文件",
                    data=output,
                    file_name="转换结果.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"处理文件时出错: {str(e)}")
            st.write("请确保上传的是有效的 Excel 文件，并且格式符合预期。")

if __name__ == "__main__":
    main()
