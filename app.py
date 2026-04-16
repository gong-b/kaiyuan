# app.py 片段修复
# 在调用 subprocess 之前明确写入文件路径
if st.button("▶️ 开始筛选", type="primary"):
    # ...
    # 强制覆盖 config.py 中预期的文件名
    with open(DATA_DIR / "new_hongji.xlsx", "wb") as f:
        f.write(new_hongji_file.getbuffer())
    with open(DATA_DIR / "last_year.xlsx", "wb") as f:
        f.write(last_year_file.getbuffer())
    
    # 执行
    # ...
