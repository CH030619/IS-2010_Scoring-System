import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import math
import re
import io

# --- 0. Access Control ---
password = st.text_input("Enter Access Code", type="password")
if password != "TEK2026":
    st.info("Please enter the correct access code.")
    st.stop() 

# --- 1. Page Setting ---
st.set_page_config(page_title="IS 2010 Scoring System", layout="wide")
st.title("IS 2010 Scoring System")

def check_logic_equivalence(prof_f, stud_f, prof_v, stud_v):
    # 1. round off error (1.999... vs 2.0)
    # 2. formula error solved (A1*2 vs 2*A1)
    try:
        v_p_float = float(prof_v) if prof_v is not None else 0.0
        v_s_float = float(stud_v) if stud_v is not None else 0.0
        values_match = math.isclose(v_p_float, v_s_float, rel_tol=1e-9, abs_tol=1e-9)
    except (ValueError, TypeError):
        # if not real numbers, compare
        values_match = str(prof_v).strip().upper() == str(stud_v).strip().upper()

    if values_match:
        # if values match, compare the formulas
        p_norm = str(prof_f).replace(" ", "").upper() if prof_f else ""
        s_norm = str(stud_f).replace(" ", "").upper() if stud_f else ""
        
        if p_norm == s_norm:
            return True
        
        if sorted(list(p_norm)) == sorted(list(s_norm)):
            return True
            
    return False

def reset_grading():
    st.session_state['grading_done'] = False
    st.session_state['summary_df'] = None
    st.session_state['wrong_df'] = None
    st.session_state['total_questions'] = 0
if 'grading_done' not in st.session_state:
    reset_grading()



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

# --- 2. File uploaders ---
col1, col2 = st.columns(2)
with col1:
    prof_file = st.file_uploader("1. Upload Professor's File", type=['xlsx'], key="prof", on_change = reset_grading)
with col2:
    student_files = st.file_uploader("2. Upload Student's File(s)", type=['xlsx'], accept_multiple_files=True, key="stud", on_change = reset_grading)

# --- 3. Grading Logic & UI Section ---
if prof_file and student_files:
    st.subheader("Selection & Execution")
    
    # [Security] Validation for misplaced files
    is_valid_upload = True
    uid_pattern = re.compile(r'[uU]\d{7}(?!\d)')

    # Check if professor's slot contains a student file
    if uid_pattern.search(prof_file.name):
        st.error(f"**Upload Error:** The file '{prof_file.name}' in the Professor's slot appears to be a student file. Please check again.")
        is_valid_upload = False
        
    # Check if student slots contain files without a valid UnID
    invalid_files = [f.name for f in student_files if not uid_pattern.search(f.name)]
    if invalid_files:
        st.error(f"**Naming Error:** Files without a valid UnID detected: {', '.join(invalid_files)}")
        st.info("Student files must include a valid ID (e.g., U1234567) to be processed.")
        is_valid_upload = False

    ui_col1, ui_col2 = st.columns([2, 1])
    with ui_col1:
        selected_name = st.selectbox(
            "Which color did you use for the answer cells?",
            ["--- Select the color you used in Excel ---"] + list(standard_colors.keys())
        )
    with ui_col2:
        st.write("") 
        st.write("")
        # Start button is disabled if validation fails
        start_button = st.button("Start Grading Process", width='stretch', disabled=not is_valid_upload)

    if not is_valid_upload:
        st.stop()

    if selected_name != "--- Select the color you used in Excel ---":
        target_hex_6 = standard_colors[selected_name].lstrip('#').upper()

        if start_button:
            # Process Professor's File
            p_bytes = prof_file.read()
            wb_p_f = load_workbook(io.BytesIO(p_bytes), data_only=False)
            wb_p_v = load_workbook(io.BytesIO(p_bytes), data_only=True)

            summary_data = []
            wrong_data = []
            check_map = {}
            total_questions = 0

            # Identify Answer Cells
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

            # Grading Students
            progress_bar = st.progress(0)
            st.warning("**System Note:** Processing valid student files. Any file without a proper UNID has been excluded from this run.")

            # --- Grading Students (Optimized) ---
            progress_bar = st.progress(0)
            
            # Start scoring
            for i, s_file in enumerate(student_files):
                try:
                    uid = uid_pattern.search(s_file.name).group().upper()
                    
                    correct_count = 0
                    s_bytes = s_file.read()
                    wb_s_f = load_workbook(io.BytesIO(s_bytes), data_only=False)
                    wb_s_v = load_workbook(io.BytesIO(s_bytes), data_only=True)

                    for sn, cells in check_map.items():
                        if sn not in wb_s_v.sheetnames: continue
                        for c in cells:
                            pf, sf = wb_p_f[sn][c].value, wb_s_f[sn][c].value
                            pv, sv = wb_p_v[sn][c].value, wb_s_v[sn][c].value
                            
                            # comparison between equation, values 
                            if check_logic_equivalence(pf, sf, pv, sv):
                                correct_count += 1
                            else:
                                wrong_data.append({
                                    "UnID": uid, "Sheet": sn, "Cell": c,
                                    "Prof Formula": pf, "Student Formula": sf,
                                    "Prof Value": pv, "Student Value": sv
                                })
                                                        
                    summary_data.append({
                        "UnID": uid, 
                        "File": s_file.name, 
                        "Score": f"{correct_count}/{total_questions}", 
                        "Raw_Score": correct_count
                    })

                except Exception as e:
                    # Recognize files with errors except file name errors
                    st.warning(f"System Error with {s_file.name}: {e}")

                finally:
                    progress_bar.progress((i + 1) / len(student_files))

            progress_bar.empty()

            # Save results to session state
            df_res = pd.DataFrame(summary_data).sort_values(by="Raw_Score", ascending=False).drop(columns=["Raw_Score"])
            st.session_state['summary_df'] = df_res
            st.session_state['wrong_df'] = pd.DataFrame(wrong_data)
            st.session_state['total_questions'] = total_questions
            st.session_state['grading_done'] = True

# --- 4. Display Results Section ---
if st.session_state['grading_done']:
    st.divider()
    st.success(f"Grading Complete! Found {st.session_state['total_questions']} answer cells per file.")

    # Summary Table
    st.subheader("Grading Summary")
    st.info("Click a row in the table below to view the individual narrative feedback report.")
    
    event = st.dataframe(
        st.session_state['summary_df'],
        width='stretch',
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
        key="summary_table"
    )

    # Error Analysis Visualization
    st.divider()
    df_wrong = st.session_state['wrong_df']
    if not df_wrong.empty:
        st.subheader("Error Analysis (Frequent Mistakes)")
        error_counts = df_wrong["Cell"].value_counts().reset_index()
        error_counts.columns = ["Cell", "Number of Errors"]
        st.bar_chart(data=error_counts, x="Cell", y="Number of Errors", color="#ff4b4b")

        # Detailed Individual Feedback
        selected_uid = None
        if event.selection.rows:
            selected_index = event.selection.rows[0]
            selected_uid = st.session_state['summary_df'].iloc[selected_index]["UnID"]

        if selected_uid:
            st.markdown(f"### Individual Feedback Report: `{selected_uid}`")
            display_df = df_wrong[df_wrong["UnID"] == selected_uid]
            
            if display_df.empty:
                st.balloons()
                st.success(f"Perfect Score! Student **{selected_uid}** correctly completed all tasks.")
            else:
                
                for idx, row in display_df.iterrows():
                    f_prof_clean = str(row['Prof Formula']).strip().upper() if row['Prof Formula'] else ""
                    f_stud_clean = str(row['Student Formula']).strip().upper() if row['Student Formula'] else ""
                    
                    is_logic_same = (f_prof_clean == f_stud_clean) and (f_prof_clean != "")

                    if is_logic_same:
                        p_ans = f"{row['Prof Formula']} ({row['Prof Value']})"
                        s_ans = f"{row['Student Formula']} ({row['Student Value']})"
                    else:
                        p_ans = row['Prof Formula'] if row['Prof Formula'] else row['Prof Value']
                        s_ans = row['Student Formula'] if row['Student Formula'] else row['Student Value']

                    with st.container(border=True):
                        st.markdown(f"**Item {idx+1} (Cell {row['Cell']} / Sheet: {row['Sheet']})**")
                        
                        c1, c2 = st.columns(2)
                        with c1:
                            st.markdown(f"**Student's Answer**\n<div style='color:#ff4b4b; font-size:1.1rem; font-weight:bold; background-color:#fff5f5; padding:8px; border-radius:5px; border-left:4px solid #ff4b4b;'>{s_ans}</div>", unsafe_allow_html=True)
                        with c2:
                            st.markdown(f"**Suggested Solution**\n<div style='color:#008000; font-size:1.1rem; font-weight:bold; background-color:#f0fff0; padding:8px; border-radius:5px; border-left:4px solid #008000;'>{p_ans}</div>", unsafe_allow_html=True)
                        if is_logic_same:
                            st.markdown(f"<div style='margin-top:10px; font-style: italic; color: #555;'>Analysis: The formulas are identical, but calculated values differ.</div>", unsafe_allow_html=True)

                        else:
                            st.markdown(f"<div style='margin-top:10px; font-style: italic; color: #555;'>Analysis: An incorrect answer was identified. Check the logic or cell references.</div>", unsafe_allow_html=True)

            st.caption("Note: Comparison includes both formulas and final calculated values.")
        else:
            st.warning("Please select a student from the Summary Table to generate their feedback report.")

        # Download Results
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
        st.success("Perfect! All submitted files scored 100%.")
else:
    if not prof_file or not student_files:
        st.info("Ready to grade. Please upload the Professor's and Student's files.")
