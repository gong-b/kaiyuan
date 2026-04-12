# app.py 中执行main.py的片段优化
with st.spinner("🔍 正在筛选邮件和处理数据..."):
    try:
        original_cwd = os.getcwd()
        os.chdir(BASE_DIR)
        
        # 执行并捕获输出
        result = subprocess.run(
            [sys.executable, "main.py"],
            capture_output=True,
            encoding="utf-8",
            errors="replace",
            timeout=600
        )
        os.chdir(original_cwd)

        # 展示日志（按级别高亮）
        st.subheader("📜 运行日志")
        log_content = result.stdout + "\n" + result.stderr
        # 拆分日志行，高亮错误/致命信息
        log_lines = log_content.split("\n")
        highlighted_log = []
        for line in log_lines:
            if any(level in line for level in ["CRITICAL", "ERROR"]):
                highlighted_log.append(f"<span style='color: red;'>{line}</span>")
            elif "WARNING" in line:
                highlighted_log.append(f"<span style='color: orange;'>{line}</span>")
            else:
                highlighted_log.append(line)
        # 折叠显示高亮日志
        with st.expander("查看完整日志（错误已标红）", expanded=True):
            st.markdown("<br>".join(highlighted_log), unsafe_allow_html=True)

        # 结果判断
        if result.returncode == 0:
            st.success("✅ 筛选完成！")
        else:
            st.error(f"❌ 筛选过程出错（返回码：{result.returncode}）")
            # 提取关键错误提示
            error_lines = [line for line in log_lines if "ERROR" in line or "CRITICAL" in line]
            if error_lines:
                st.subheader("🔴 关键错误信息")
                for line in error_lines[:10]:  # 只显示前10条
                    st.code(line)

    except subprocess.TimeoutExpired:
        st.error("❌ 处理超时（超过10分钟）！请检查：\n1. 邮件数量是否过多\n2. 邮箱连接是否稳定\n3. 日期范围是否过大")
    except Exception as e:
        st.error(f"❌ 执行失败：{str(e)}")
        st.code(f"详细错误：{str(e)}", language="text")
