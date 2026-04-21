for idx, (uid, msg) in enumerate(mails):
    try:
        # 1. 过滤自己发送/回复/转发邮件
        sender_email = parseaddr(msg.get("From", ""))[1]
        if sender_email == user:
            bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封（跳过自己发送的邮件）")
            continue

        subj = ep.parse_subject(msg)
        if any(prefix in subj[:5] for prefix in ["RE:", "FW:", "回复:", "转发:"]):
            bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封（跳过回复/转发邮件）")
            continue

        # 日期过滤（增加异常捕获）
        try:
            d_utc = parsedate_to_datetime(msg["Date"])
            d_local = d_utc.astimezone()
            if not (s_date <= d_local.date() <= e_date):
                bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封（跳过非目标日期邮件）")
                continue
        except Exception as e:
            st.warning(f"邮件{uid}日期解析失败：{str(e)}，跳过")
            continue

        # 2. 修复附件提取逻辑：兼容直接发送的带附件邮件
        raw_msg = msg
        # 优先处理嵌套邮件（会话），否则用原始邮件
        if msg.is_multipart():
            has_rfc822 = False
            for part in msg.walk():
                if part.get_content_type() == "message/rfc822":
                    raw_msg = message_from_bytes(part.get_payload(decode=True))
                    has_rfc822 = True
                    break
            # 若没有嵌套邮件，直接用原始msg解析附件
            if not has_rfc822:
                raw_msg = msg

        # 3. 解析附件 + 提取报名班级（增强容错）
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            # 调试：打印附件提取前的邮件类型
            st.write(f"调试-邮件{uid}：是否多部分={raw_msg.is_multipart()}，发件人={sender_email}")
            
            # 修复附件提取：确保EmailParser的extract_docx_attachments能处理普通邮件
            docs = ep.extract_docx_attachments(raw_msg, tmp_path)
            f_name = "未知"
            f_sid = ""
            apply_class = "未知班级"

            # 3.1 附件提取失败：从主题提取
            if not docs:
                st.write(f"调试-邮件{uid}：无docx附件，主题={subj}")  # 调试用
                pattern = re.search(r"([^+]+)\+(\d{8,10})\+(.*?班)", subj)
                if pattern:
                    f_name = pattern.group(1).strip()
                    f_sid = pattern.group(2).strip()
                    apply_class = pattern.group(3).strip()
                # 无附件记录
                current_record = {
                    "name": f_name,
                    "sid": f_sid,
                    "class": apply_class,
                    "status": "reject",
                    "reason": "缺失DOCX附件",
                    "subject": subj,
                    "date": d_local
                }
            else:
                # 3.2 附件提取成功：解析信息
                st.write(f"调试-邮件{uid}：找到附件{docs}")  # 调试用
                info = dp.parse(str(docs[0]))
                f_name = info.get("name", "未知")
                f_sid = info.get("sid", "")
                
                # 优先从附件提班级，无则从主题提（增强正则）
                apply_class = info.get("apply_class", "")
                if not apply_class:
                    class_match = re.search(r"([^+]+班)", subj)  # 放宽正则匹配
                    apply_class = class_match.group(1).strip() if class_match else "未知班级"

                if not f_sid:
                    current_record = {
                        "name": f_name,
                        "sid": "",
                        "class": apply_class,
                        "status": "reject",
                        "reason": "附件内无学号",
                        "subject": subj,
                        "date": d_local
                    }
                else:
                    # 核心：首次录取班级逻辑
                    if f_sid in student_admitted_class:
                        admitted_class = student_admitted_class[f_sid]
                        if apply_class == admitted_class:
                            current_record = {
                                "name": f_name,
                                "sid": f_sid,
                                "class": apply_class,
                                "status": "accept",
                                "reason": "",
                                "remark": f"审核通过（已录取{admitted_class}）",
                                "date": d_local
                            }
                        else:
                            current_record = {
                                "name": f_name,
                                "sid": f_sid,
                                "class": apply_class,
                                "status": "reject",
                                "reason": f"重复报名（已录取{admitted_class}，本次报名{apply_class}）",
                                "subject": subj,
                                "date": d_local
                            }
                    else:
                        # 未录取过：按规则审核
                        if f_sid in B:
                            current_record = {
                                "name": f_name, "sid": f_sid, "class": apply_class,
                                "status": "reject", "reason": "黑名单人员",
                                "subject": subj, "date": d_local
                            }
                        elif f_sid in H:
                            current_record = {
                                "name": f_name, "sid": f_sid, "class": apply_class,
                                "status": "accept", "reason": "",
                                "remark": f"新鸿基(录取{apply_class})", "date": d_local
                            }
                            student_admitted_class[f_sid] = apply_class
                        elif f_sid in L:
                            current_record = {
                                "name": f_name, "sid": f_sid, "class": apply_class,
                                "status": "reject", "reason": "去年已录取",
                                "subject": subj, "date": d_local
                            }
                        elif not info.get("is_supported", False):
                            current_record = {
                                "name": f_name, "sid": f_sid, "class": apply_class,
                                "status": "reject", "reason": "非资助对象",
                                "subject": subj, "date": d_local
                            }
                        elif info.get("reason_length", 0) < Config.MIN_REASON_LENGTH:
                            current_record = {
                                "name": f_name, "sid": f_sid, "class": apply_class,
                                "status": "reject", "reason": f"理由不足({info['reason_length']}字)",
                                "subject": subj, "date": d_local
                            }
                        else:
                            current_record = {
                                "name": f_name, "sid": f_sid, "class": apply_class,
                                "status": "accept", "reason": "",
                                "remark": f"审核通过（录取{apply_class}）", "date": d_local
                            }
                            student_admitted_class[f_sid] = apply_class

            # ========== 去重逻辑：保留最优记录 ==========
            if f_sid and f_sid != "未知":
                if f_sid not in student_records:
                    student_records[f_sid] = current_record
                else:
                    existing = student_records[f_sid]
                    # 录取优先 + 同状态取最新
                    if existing["status"] == "reject" and current_record["status"] == "accept":
                        student_records[f_sid] = current_record
                    elif existing["status"] == current_record["status"]:
                        if current_record["date"] > existing["date"]:
                            student_records[f_sid] = current_record
            else:
                # 无有效学号：加入拒绝列表（不再被覆盖）
                no_final.append({
                    "学号": f_sid if f_sid else "未知",
                    "姓名": f_name,
                    "报名班级": apply_class,
                    "原因": current_record["reason"],
                    "原主题": subj,
                    "报名时间": d_local.strftime("%Y-%m-%d %H:%M")
                })

        bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封")

    except Exception as e:
        err_msg = f"解析异常: {str(e)[:50]}"  # 加长异常信息
        st.error(f"邮件{uid}解析失败：{err_msg}")  # 打印具体异常
        no_final.append({
            "学号": "?", "姓名": "?", "报名班级": "未知",
            "原因": err_msg, "原主题": "",
            "报名时间": ""
        })
        bar.progress((idx+1)/total, text=f"解析中：{idx+1}/{total} 封（异常）")
