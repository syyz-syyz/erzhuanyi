import streamlit as st
import pandas as pd
import io

def convert_2d_to_1d(df, fixed_columns):
    """
    将二维 DataFrame 转换为一维格式
    用户可指定固定列，其余列将被转换为字段名和对应值
    """
    if df.empty or not fixed_columns or len(fixed_columns) >= len(df.columns):
        return pd.DataFrame()
    
    # 创建一个空的 DataFrame 用于存储结果
    result_df = pd.DataFrame(columns=list(fixed_columns) + ['字段名称', '值内容'])
    
    # 计算总工作量
    total_work = (len(df.columns) - len(fixed_columns)) * len(df)
    current_work = 0
    
    # 创建进度条
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    # 获取转换列的列名
    convert_columns = [col for col in df.columns if col not in fixed_columns]
    
    # 遍历每一个需要转换的字段
    for field_name in convert_columns:
        # 遍历每一行
        for _, row in df.iterrows():
            # 获取固定列的值
            fixed_values = [row[col] for col in fixed_columns]
            # 获取当前字段的值
            value = row[field_name]
            
            # 添加到结果 DataFrame
            new_row = pd.DataFrame(
                [fixed_values + [field_name, value]],
                columns=result_df.columns
            )
            result_df = pd.concat([result_df, new_row], ignore_index=True)
            
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
            # 获取所有表名
            xls = pd.ExcelFile(uploaded_file)
            sheet_names = xls.sheet_names
            
            # 选择工作表
            selected_sheet = st.selectbox(
                "选择要处理的工作表",
                options=sheet_names,
                index=0
            )
            
            # 读取选中的工作表
            df = xls.parse(selected_sheet)
            
            # 显示原始数据预览
            st.subheader(f"工作表 '{selected_sheet}' 的数据预览")
            st.dataframe(df.head())
            
            # 检查是否有足够的列进行转换
            if len(df.columns) < 2:
                st.error("Excel 文件至少需要两列才能进行转换。")
                return
                
            # 用户选择固定列
            fixed_columns = st.multiselect(
                "选择固定列（这些列将保持不变）",
                options=df.columns.tolist(),
                default=[df.columns[0]] if len(df.columns) > 0 else [],
                help="这些列将作为固定列，其余列将被转换为字段名和值"
            )
            
            # 确保至少选择了一列
            if not fixed_columns:
                st.error("请至少选择一列作为固定列。")
                return
            
            # 显示固定列和转换列的预览
            st.subheader("列配置预览")
            convert_cols = [col for col in df.columns if col not in fixed_columns]
            st.write(f"**固定列** ({len(fixed_columns)}): {', '.join(fixed_columns)}")
            st.write(f"**转换列** ({len(convert_cols)}): {', '.join(convert_cols)}")
            
            # 执行转换
            if st.button("开始转换"):
                st.subheader("正在转换数据...")
                converted_df = convert_2d_to_1d(df, fixed_columns)
                
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
                        file_name=f"{selected_sheet}_转换结果.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        except Exception as e:
            st.error(f"处理文件时出错: {str(e)}")
            st.write("请确保上传的是有效的 Excel 文件，并且格式符合预期。")

if __name__ == "__main__":
    main()    
