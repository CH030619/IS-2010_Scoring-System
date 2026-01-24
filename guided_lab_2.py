import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import re
import io

password = st.text_input("Enter Access Code", type="password")
if password != "TEK2026":
    st.info("Please enter the correct access code.")
    st.stop() 

# 1. Page Setting
st.set_page_config(page_title="IS 2010 Scoring System", layout="wide")
st.title("IS 2010 Scoring System")

# Initialize session state to store grading results
if 'grading_done' not in st.session_state:
    st.session_state['grading_done'] = False
    st.session_state['summary_df'] = None
    st.session_state['wrong_df'] = None
    st.session_state['total_questions'] = 0

# --- Step 1: Professor's Color Guide ---
with st.expander("IMPORTANT: Professor's Color Guide", expanded=True):
    st.info("To ensure accurate grading, please use one of these **10 Standard Colors** for your answer cells in Excel.")
    
    standard_colors = {
        "Dark Red": "#C00000", "Red": "#FF0000", "Orange": "#FFC000", 
        "Yellow": "#FFFF00", "Light Green": "#92D050", "Green": "#00B050", 
        "Light Blue": "#00B0F0", "Blue": "#0070C0", "Dark Blue": "#002060", "Purple": "#7030A0"
    }
    
    cols = st.columns(10)
    for idx, (name, hex_val) in enumerate(standard_colors.items()):
        with cols[idx]:
            st.markdown(f"<div style='background-color:{hex_val}; width:100%; height:30px; border-radius:3px; border:1px solid gray;'></div>", unsafe_allow_html=True)
            st.caption(f"**{name}**")

st.divider()

# 2. File uploaders
col1, col2 = st.columns(2)
with col1:
    prof_file = st.file_uploader("1. Upload Professor's File", type=['xlsx'], key="prof")
with col2:
    student_files = st.file_uploader("2. Upload Student's File(s)", type=['xlsx'], accept_multiple_files=True, key="stud")

# 3. Grading Logic & UI Section
if prof_file and student_files:
    st.subheader("Selection & Execution")
    
    ui_col1, ui_col2 = st.columns([2, 1])
    with ui_col1:
        selected_name = st.selectbox(
            "Which color did you use for the answer cells?",
            ["--- Select the color you used in Excel ---"] + list(standard_colors.keys())
        )
    with ui_col2:
        st.write("") # Spacing
        st.write("")
        start_button = st.button("Start Grading Process", width='stretch')

    if selected_name != "--- Select the color you used in Excel ---":
        target_hex_6 = standard_colors[selected_name].lstrip('#').upper()

        if start_button:
            # Process files
            p_bytes = prof_file.read()
            wb_p_f = load_workbook(io.BytesIO(p_bytes), data_only=False)
            wb_p_v = load_workbook(io.BytesIO(p_bytes), data_only=True)

            summary_data = []
            wrong_data = []
            check_map = {}
            total_questions = 0

            # --- Identify Answer Cells ---
            for sn in wb_p_f.sheetnames:
                sheet = wb_p_f[sn]
                cells = []
                for row in sheet.iter_rows():
                    for c in row:
                        if c.fill and c.fill.fill_type == 'solid':
                            raw_color = str(c.fill.start_color.rgb).upper()
                            clean_color = raw_color[-6:] if len(raw_color) >= 6 else raw_color
                            if clean_color == target_hex_6:
                                cells.append(c.coordinate)
                if cells:
                    check_map[sn] = cells
                    total_questions += len(cells)

            if not check_map:
                st.error(f"Error: No cells found with color '{selected_name}'.")
                st.stop()

            # --- Grading Students ---
            progress_bar = st.progress(0)
            uid_pattern = re.compile(r'[uU]\d{7}')

            for i, s_file in enumerate(student_files):
                try:
                    correct_count = 0
                    s_bytes = s_file.read()
                    wb_s_f = load_workbook(io.BytesIO(s_bytes), data_only=False)
                    wb_s_v = load_workbook(io.BytesIO(s_bytes), data_only=True)
                    
                    match = uid_pattern.search(s_file.name)
                    uid = match.group() if match else s_file.name

                    for sn, cells in check_map.items():
                        if sn not in wb_s_v.sheetnames: continue
                        for c in cells:
                            pf, sf = wb_p_f[sn][c].value, wb_s_f[sn][c].value
                            pv, sv = wb_p_v[sn][c].value, wb_s_v[sn][c].value
                            
                            f_p, f_s = str(pf).strip() if pf else "", str(sf).strip() if sf else ""
                            v_p, v_s = str(pv).strip() if pv else "", str(sv).strip() if sv else ""

                            if f_p.upper() == f_s.upper() and v_p.upper() == v_s.upper():
                                correct_count += 1
                            else:
                                wrong_data.append({
                                    "UnID": uid, "Sheet": sn, "Cell": c,
                                    "Prof Formula": pf, "Student Formula": sf,
                                    "Prof Value": pv, "Student Value": sv
                                })
                    
                    summary_data.append({"UnID": uid, "File": s_file.name, "Score": f"{correct_count}/{total_questions}", "Raw_Score": correct_count})
                    progress_bar.progress((i + 1) / len(student_files))
                except Exception as e:
                    st.warning(f"Error processing {s_file.name}: {e}")

            progress_bar.empty()

            # Store in session state
            df_res = pd.DataFrame(summary_data).sort_values(by="Raw_Score", ascending=False).drop(columns=["Raw_Score"])
            st.session_state['summary_df'] = df_res
            st.session_state['wrong_df'] = pd.DataFrame(wrong_data)
            st.session_state['total_questions'] = total_questions
            st.session_state['grading_done'] = True

# --- Display Results Section ---
if st.session_state['grading_done']:
    st.divider()
    st.success(f"Grading Complete! Found {st.session_state['total_questions']} answer cells.")

    # 1. Summary Table with Selection
    st.subheader("Grading Summary")
    st.info("**Instructions:** Click a row in the table below to generate a detailed feedback report for that specific student.")
    
    event = st.dataframe(
        st.session_state['summary_df'],
        width='stretch',
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
        key="summary_table"
    )

    # 2. Statistics & Visualization
    st.divider()
    st.subheader("Error Analysis")
    
    df_wrong = st.session_state['wrong_df']
    if not df_wrong.empty:
        # Chart: Error Frequency per Cell
        error_counts = df_wrong["Cell"].value_counts().reset_index()
        error_counts.columns = ["Cell", "Number of Errors"]
        st.bar_chart(data=error_counts, x="Cell", y="Number of Errors", color="#ff4b4b")

        # Filtering based on table selection
        selected_uid = None
        if event.selection.rows:
            selected_index = event.selection.rows[0]
            selected_uid = st.session_state['summary_df'].iloc[selected_index]["UnID"]

        # --- [NEW] 3. Detailed Feedback Report (Formal Narrative) ---
        if selected_uid:
            st.markdown(f"### Individual Feedback Report")
            st.markdown(f"**Target Student ID:** `{selected_uid}`")
            
            display_df = df_wrong[df_wrong["UnID"] == selected_uid]
            
            if display_df.empty:
                st.success(f"**Final Evaluation:** Student **{selected_uid}** demonstrated a perfect understanding of the tasks. No errors were found in any of the required formulas or values.")
            else:
                st.markdown("---")
                # Making a short report
                # --- [REFINED] Individual Feedback Report Design ---
                for idx, row in display_df.iterrows():
                    p_ans = row['Prof Formula'] if row['Prof Formula'] else row['Prof Value']
                    s_ans = row['Student Formula'] if row['Student Formula'] else row['Student Value']
                    
                    # Report layout
                    with st.container(border=True):
                        # Item, cell, rows #
                        st.markdown(f"Item {idx+1} (Cell {row['Cell']} / {row['Sheet']})")
                        
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown(f"**Student's Answer**\n<div style='color:#ff4b4b; font-size:1.1rem; font-weight:bold; background-color:#fff5f5; padding:8px; border-radius:5px; border-left:4px solid #ff4b4b;'>{s_ans}</div>", unsafe_allow_html=True)
                        with c2:
                            st.markdown(f"**Suggested Solution**\n<div style='color:#008000; font-size:1.1rem; font-weight:bold; background-color:#f0fff0; padding:8px; border-radius:5px; border-left:4px solid #008000;'>{p_ans}</div>", unsafe_allow_html=True)
                        
                        # Specific feedback
                        st.markdown(
                            f"<div style='margin-top:15px; padding-top:10px; border-top:1px solid #eee;'>"
                            f"<strong>Analysis:</strong> Incorrect answer was identified in this cell. This suggests a potential error in the calculation logic or a misinterpretation of the task requirements. "
                            f"Please re-examine the instructions to ensure the formula aligns with the intended output. "
                            f"</div>", 
                            unsafe_allow_html=True
                        )
                st.caption("Note: This report is generated based on a comparison between the student's submission and the professor's file.")
        else:
            st.warning("Please select a student from the table above to generate the narrative feedback report.")

        # 4. Download Button
        st.divider()
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            st.session_state['summary_df'].to_excel(writer, index=False, sheet_name="Summary")
            df_wrong.to_excel(writer, index=False, sheet_name="All_Errors")
        
        st.download_button(
            label="Download Full Grading Report (.xlsx)",
            data=output.getvalue(),
            file_name=f"Grading_Result_{selected_name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch'
        )
    else:
        st.balloons()
        st.success("Perfect! All students scored 100%.")

else:
    if not prof_file or not student_files:
        st.info("Please upload the Professor's guide and Student's files to initiate the process.")
