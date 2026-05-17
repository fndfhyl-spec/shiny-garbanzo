"""
PDF合并工具
支持上传多个PDF文件，自定义合并顺序，一键合并下载
"""

import streamlit as st
from PyPDF2 import PdfMerger
import io

# 页面配置
st.set_page_config(
    page_title="PDF合并工具",
    page_icon="📄",
    layout="wide"
)

# 标题
st.title("📄 PDF合并工具")
st.markdown("上传多个PDF文件，拖拽调整顺序，一键合并下载")

st.markdown("---")

# 上传PDF文件
uploaded_files = st.file_uploader(
    "上传PDF文件",
    type=['pdf'],
    accept_multiple_files=True,
    help="可同时选择多个PDF文件"
)

if uploaded_files:
    st.success(f"已上传 {len(uploaded_files)} 个文件")
    
    # 显示文件列表，允许调整顺序
    st.markdown("### 📋 文件顺序（可拖拽调整）")
    
    # 使用session_state保存顺序
    if 'file_order' not in st.session_state or len(st.session_state.file_order) != len(uploaded_files):
        st.session_state.file_order = list(range(len(uploaded_files)))
    
    # 显示每个文件，带上下移动按钮
    cols_per_row = 3
    for i in range(0, len(uploaded_files), cols_per_row):
        cols = st.columns(cols_per_row)
        for j, col in enumerate(cols):
            idx = i + j
            if idx < len(uploaded_files):
                file_idx = st.session_state.file_order[idx]
                file = uploaded_files[file_idx]
                
                with col:
                    # 文件信息容器
                    with st.container(border=True):
                        st.markdown(f"**{idx + 1}.** {file.name}")
                        
                        # 上下移动按钮
                        btn_cols = st.columns(2)
                        with btn_cols[0]:
                            if idx > 0 and st.button("⬆️ 上移", key=f"up_{idx}"):
                                # 交换顺序
                                st.session_state.file_order[idx], st.session_state.file_order[idx-1] = \
                                    st.session_state.file_order[idx-1], st.session_state.file_order[idx]
                                st.rerun()
                        
                        with btn_cols[1]:
                            if idx < len(uploaded_files) - 1 and st.button("⬇️ 下移", key=f"down_{idx}"):
                                # 交换顺序
                                st.session_state.file_order[idx], st.session_state.file_order[idx+1] = \
                                    st.session_state.file_order[idx+1], st.session_state.file_order[idx]
                                st.rerun()
    
    st.markdown("---")
    
    # 显示合并预览
    st.markdown("### 📝 合并预览")
    preview_text = " → ".join([uploaded_files[st.session_state.file_order[i]].name 
                               for i in range(len(uploaded_files))])
    st.info(preview_text)
    
    # 合并按钮
    if st.button("🔗 合并PDF", type="primary"):
        with st.spinner("正在合并..."):
            try:
                merger = PdfMerger()
                
                # 按顺序添加文件
                for i in range(len(uploaded_files)):
                    file_idx = st.session_state.file_order[i]
                    uploaded_files[file_idx].seek(0)  # 重置文件指针
                    merger.append(uploaded_files[file_idx])
                
                # 写入内存
                output = io.BytesIO()
                merger.write(output)
                merger.close()
                output.seek(0)
                
                st.success("✅ 合并成功！")
                
                # 下载按钮
                st.download_button(
                    label="📥 下载合并后的PDF",
                    data=output,
                    file_name="merged.pdf",
                    mime="application/pdf"
                )
                
            except Exception as e:
                st.error(f"合并失败: {e}")

else:
    st.info("👆 请上传PDF文件")

# 使用说明
with st.expander("📖 使用说明"):
    st.markdown("""
    1. 点击上方上传区域，选择多个PDF文件
    2. 使用 **⬆️ 上移** 和 **⬇️ 下移** 按钮调整顺序
    3. 点击 **合并PDF** 按钮
    4. 下载合并后的文件
    """)
