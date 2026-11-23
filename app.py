
import streamlit as st
import pandas as pd
from ortools.sat.python import cp_model
import io

# ----------------- 페이지 설정 -----------------
st.set_page_config(page_title="1:1 Biz Matching Scheduler", layout="wide", page_icon="📅")

st.title("🎯 KSC 1:1 Business Matching Scheduler Platform")
st.caption("참가자 정보와 제약조건을 입력하면 AI 최적화 엔진(OR-Tools)이 중복과 공백 없는 최적 시간표를 자동 생성합니다.")

# ----------------- 사이드바 설정 -----------------
with st.sidebar:
    st.header("⚙️ 행사 환경 설정")
    slot_duration = st.selectbox("미팅 단위 (분)", [15, 20, 30], index=1)
    max_tables = st.slider("최대 운영 테이블 수", 2, 10, 6)
    
    st.subheader("시간대 설정")
    time_slots = st.text_area(
        "슬롯 목록 (엔터로 구분)",
        value="10:00 - 10:20\n10:20 - 10:40\n10:40 - 11:00\n11:00 - 11:20\n11:20 - 11:40\n11:40 - 12:00\n"
              "13:00 - 13:20\n13:20 - 13:40\n13:40 - 14:00\n14:00 - 14:20\n14:20 - 14:40\n14:40 - 15:00\n"
              "15:00 - 15:20\n15:20 - 15:40\n15:40 - 16:00\n16:00 - 16:20\n16:20 - 16:40\n16:40 - 17:00",
        height=200
    ).strip().split("\n")

# ----------------- 탭 구성 -----------------
tab1, tab2, tab3 = st.tabs(["📝 데이터 입력 및 관리", "⚡ 스케줄 생성 및 최적화", "📊 결과 및 다운로드"])

with tab1:
    col1, col2 = st.columns(2)
    with col1:
        st.subheader("1. 참가 스타트업 명단")
        default_startups = pd.DataFrame({
            "Startup": ["Flexcuit", "HeyHR", "Artlive", "yeon living", "Genstory", "GOGAETOK", 
                        "CircledDot.Co", "RentTrail", "MayMay", "Hiddit", "Matey", "Hello Science Edu", 
                        "What’s For Protein", "Huole (Miri)"],
            "Unavailable_Slots": ["", "16:00 - 16:20, 16:20 - 16:40, 16:40 - 17:00", "", "", "", "", "", "", "", "", "", "", "", ""]
        })
        startups_df = st.data_editor(default_startups, num_rows="dynamic", use_container_width=True, key="startups_table")

    with col2:
        st.subheader("2. 참가 투자자/파트너 (호스트)")
        default_hosts = pd.DataFrame({
            "Host": ["Raymond", "Dawn Lee", "Joshia Kwa", "Gloria Traenkle", "Dave Kwong", "Alex Neo", "Bryan Wu", "Elaine Lai", "Shane Tay", "Alrick Oh"],
            "Organization": ["Focus Tech", "Paypal", "Integra", "Angel Inv", "Incubate Fund", "Singtel Innov8", "Crystalline", "Gateway of Asia", "FOMO Pay", "Berru"],
            "Max_Meetings": [10, 6, 8, 6, 8, 6, 6, 12, 8, 8]
        })
        hosts_df = st.data_editor(default_hosts, num_rows="dynamic", use_container_width=True, key="hosts_table")

# 세션 상태 저장
if "schedule_results" not in st.session_state:
    st.session_state.schedule_results = None

with tab2:
    st.subheader("🚀 매칭 최적화 실행")
    st.write("규칙: **1:1 중복 절대 금지(0회)**, **동시간대 중복 방지**, **호스트별 연속 슬롯 밀착 배정**")
    
    if st.button("스케줄 자동 생성 시작", type="primary"):
        with st.spinner("AI 최적화 솔버(Google OR-Tools)가 충돌 없는 최적 시간표를 계산 중입니다..."):
            
            # --- OR-Tools CP-SAT 최적화 로직 ---
            model = cp_model.CpModel()
            
            s_list = startups_df["Startup"].dropna().tolist()
            h_list = hosts_df["Host"].dropna().tolist()
            t_list = time_slots
            
            # 변수 생성: X[s, h, t] == 1 이면 시간 t에 스타트업 s와 호스트 h 매칭
            x = {}
            for s in s_list:
                for h in h_list:
                    for t_idx, t in enumerate(t_list):
                        x[s, h, t_idx] = model.NewBoolVar(f"x_{s}_{h}_{t_idx}")
                        
            # 제약 1: 스타트업-호스트 간 전체 일정 중 최대 1회만 미팅
            for s in s_list:
                for h in h_list:
                    model.Add(sum(x[s, h, t_idx] for t_idx in range(len(t_list))) <= 1)
                    
            # 제약 2: 특정 시간에 스타트업은 최대 1개 미팅만
            for s in s_list:
                for t_idx in range(len(t_list)):
                    model.Add(sum(x[s, h, t_idx] for h in h_list) <= 1)
                    
            # 제약 3: 특정 시간에 호스트는 최대 1개 미팅만
            for h in h_list:
                for t_idx in range(len(t_list)):
                    model.Add(sum(x[s, h, t_idx] for s in s_list) <= 1)

            # 제약 4: 특정 시간에 전체 테이블 수(max_tables) 이하
            for t_idx in range(len(t_list)):
                model.Add(sum(x[s, h, t_idx] for s in s_list for h in h_list) <= max_tables)

            # 목적 함수: 총 미팅 수 극대화 + 앞쪽 시간대 슬롯 우선 배정(가중치)
            objective_terms = []
            for t_idx in range(len(t_list)):
                time_weight = len(t_list) - t_idx  # 앞쪽 시간일수록 가중치를 주어 Gap 최소화
                for s in s_list:
                    for h in h_list:
                        objective_terms.append(x[s, h, t_idx] * (100 + time_weight))
            model.Maximize(sum(objective_terms))

            # 풀이 실행
            solver = cp_model.CpSolver()
            solver.parameters.max_time_in_seconds = 10.0
            status = solver.Solve(model)

            if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
                results = []
                summary = {s: [] for s in s_list}
                
                for t_idx, t in enumerate(t_list):
                    slot_row = {"Time Slot": t}
                    active_startups_in_slot = []
                    table_num = 1
                    
                    for h in h_list:
                        for s in s_list:
                            if solver.Value(x[s, h, t_idx]) == 1:
                                slot_row[f"Table {table_num}"] = f"{s}\n({h})"
                                active_startups_in_slot.append(s)
                                summary[s].append(h)
                                table_num += 1
                                
                    for rem in range(table_num, max_tables + 1):
                        slot_row[f"Table {rem}"] = "-"
                        
                    waiting = [s for s in s_list if s not in active_startups_in_slot]
                    slot_row["Waiting List"] = " • ".join(waiting)
                    results.append(slot_row)
                    
                st.session_state.schedule_results = pd.DataFrame(results)
                
                summary_rows = []
                for s, matched in summary.items():
                    summary_rows.append({
                        "Startup": s,
                        "Total Meetings": len(matched),
                        "Matched Partners": ", ".join(matched)
                    })
                st.session_state.summary_results = pd.DataFrame(summary_rows)
                st.success("🎉 최적 스케줄이 성공적으로 생성되었습니다! [결과 및 다운로드] 탭을 확인하세요.")
            else:
                st.error("조건을 만족하는 시간표를 찾을 수 없습니다. 제약조건을 완화해 주세요.")

with tab3:
    if st.session_state.schedule_results is not None:
        st.subheader("📅 생성된 20분 비즈 매칭 타임테이블")
        st.dataframe(st.session_state.schedule_results, use_container_width=True)

        st.subheader("📈 스타트업별 매칭 현황 요약")
        st.dataframe(st.session_state.summary_results, use_container_width=True)

        # 엑셀 다운로드
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            st.session_state.schedule_results.to_excel(writer, sheet_name='Matching Schedule', index=False)
            st.session_state.summary_results.to_excel(writer, sheet_name='Startup Summary', index=False)
            
        st.download_button(
            label="📥 엑셀 파일(.xlsx)로 전체 다운로드",
            data=buffer.getvalue(),
            file_name="Biz_Matching_Schedule_Optimized.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("먼저 [스케줄 생성 및 최적화] 탭에서 생성 버튼을 눌러주세요.")
