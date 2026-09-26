import streamlit as st
import pandas as pd
import random
import io
import docx
import altair as alt
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

st.set_page_config(page_title="CU Exam Duty Roster Generator", layout="wide")

# --- CUSTOM CSS BACKGROUND & STYLING (LIGHT GREEN THEME) ---
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(135deg, #f0f7f4 0%, #d8e9e2 100%);
        background-attachment: fixed;
    }
    [data-testid="stSidebar"] {
        background-color: #eaf4ed;
        border-right: 1px solid #d0e1d4;
    }
    h1, h2, h3 {
        color: #2b583f;
        font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif;
    }
    div.stButton > button {
        background: linear-gradient(135deg, #2b583f 0%, #438a5e 100%);
        color: white;
        border-radius: 8px;
        border: none;
        padding: 0.5rem 1rem;
        font-weight: 600;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        transition: all 0.3s ease;
    }
    div.stButton > button:hover {
        background: linear-gradient(135deg, #438a5e 0%, #2b583f 100%);
        box-shadow: 0 6px 8px rgba(0,0,0,0.15);
        transform: translateY(-1px);
    }
</style>
""", unsafe_allow_html=True)


# --- HEADER WITH LEFT LOGO & FULLY CENTERED COLLEGE INFO ---
col_logo, col_info = st.columns([1, 5])

with col_logo:
    try:
        st.image("logo_pjc.png", width=110)
    except Exception:
        pass

with col_info:
    st.markdown("""
        <div style='text-align: center; margin-right: 160px;'>
            <h1 style='color: #2b583f; margin-bottom: 0px; font-size: 28px;'>🎓 University Examination Duty Roster Generator</h1>
            <h2 style='color: #3a5a48; font-size: 22px; margin-top: 2px; margin-bottom: 0px;'>Prabhu Jagatbandhu College</h2>
            <p style='color: #4a6b5d; font-size: 14px; margin-top: 0px;'>Andul-Mouri, Howrah - 711302</p>
        </div>
    """, unsafe_allow_html=True)

st.markdown("<hr style='margin-top: 5px; margin-bottom: 25px; border: 1px solid #d0e1d4;'>", unsafe_allow_html=True)

# --- SIDEBAR: Guide & Instructions + Upload ---
st.sidebar.header("📖 User Guide & Instructions")
st.sidebar.markdown("""
1. **Upload CSV**: Upload your teacher availability dataset via the file uploader below.
2. **Add Dates**: Pick examination dates on the main screen and click **Add Exam Date**.
3. **Set Quotas**: Configure session schedules (*Both Halves*, *1st Half Only*, *2nd Half Only*) and invigilator counts.
4. **Manage Opt-Ins & Leaves**: Use the GL opt-in tool, emergency leave selector, and half-day preferences.
5. **Generate & Download**: Click **Generate Duty Roster** to view results, download both the Word document and Excel duty summary file seamlessly.
""")

st.sidebar.markdown("---")
st.sidebar.header("1. Teacher Data Upload")
uploaded_file = st.sidebar.file_uploader("Upload Teacher Availability CSV", type=["csv"])

# --- MAIN APP INTERFACE ---
if uploaded_file is not None:
    teachers_df = pd.read_csv(uploaded_file)
    all_teachers = sorted(teachers_df["Teacher/Entry"].unique().tolist())

    st.markdown("---")
    
    # --- NON-CONTINUOUS EXAM DATES MANAGER ---
    st.subheader("📅 Examination Dates & Session Schedule Manager")
    st.markdown("Add individual examination dates. Then configure which halves are active and how many invigilators are required.")
    
    col_date_input, col_date_list = st.columns([1, 1])
    
    with col_date_input:
        if "exam_dates_list" not in st.session_state:
            st.session_state.exam_dates_list = []
            
        selected_single_date = st.date_input("Select Exam Date")
        
        if st.button("➕ Add Exam Date"):
            date_str = selected_single_date.strftime("%Y-%m-%d")
            if date_str not in st.session_state.exam_dates_list:
                st.session_state.exam_dates_list.append(date_str)
                st.session_state.exam_dates_list.sort()
                st.success(f"Added exam date: {date_str}")
            else:
                st.warning(f"Date {date_str} is already in the list.")

    with col_date_list:
        st.write("**Current Scheduled Exam Dates:**")
        if st.session_state.exam_dates_list:
            for idx, d_str in enumerate(st.session_state.exam_dates_list):
                c1, c2 = st.columns([3, 1])
                c1.text(f"• {d_str} ({pd.to_datetime(d_str).strftime('%A')})")
                if c2.button("❌ Remove", key=f"rm_date_{idx}"):
                    st.session_state.exam_dates_list.pop(idx)
                    if "req_data" in st.session_state and d_str in st.session_state.req_data:
                        del st.session_state.req_data[d_str]
                    st.session_state.roster_generated = False
                    st.rerun()
            
            if st.button("Clear All Dates"):
                st.session_state.exam_dates_list = []
                st.session_state.req_data = {}
                st.session_state.roster_generated = False
                st.rerun()
        else:
            st.info("No exam dates added yet. Pick a date and click 'Add Exam Date'.")

    # --- DYNAMIC REQUIREMENTS & HALF SELECTION TABLE ---
    if st.session_state.exam_dates_list:
        st.markdown("### 🔢 Configure Active Halves & Quotas per Date")
        st.markdown("Choose whether each date has **Both Halves**, **1st Half Only**, or **2nd Half Only**, and set the invigilator counts.")
        
        if "req_data" not in st.session_state:
            st.session_state.req_data = {}
            
        req_rows = []
        for d_str in st.session_state.exam_dates_list:
            day_name = pd.to_datetime(d_str).strftime('%A')
            existing = st.session_state.req_data.get(d_str, {
                "Session Schedule": "Both Halves", 
                "1st Half Required": 5, 
                "2nd Half Required": 5
            })
            req_rows.append({
                "Date": d_str,
                "Day": day_name,
                "Session Schedule": existing.get("Session Schedule", "Both Halves"),
                "1st Half Required": existing.get("1st Half Required", 5),
                "2nd Half Required": existing.get("2nd Half Required", 5)
            })
            
        req_df = pd.DataFrame(req_rows)
        
        edited_req_df = st.data_editor(
            req_df,
            column_config={
                "Date": st.column_config.TextColumn("Date", disabled=True),
                "Day": st.column_config.TextColumn("Day", disabled=True),
                "Session Schedule": st.column_config.SelectboxColumn(
                    "Session Schedule",
                    options=["Both Halves", "1st Half Only", "2nd Half Only"],
                    required=True
                ),
                "1st Half Required": st.column_config.NumberColumn("1st Half Count", min_value=0, step=1, required=True),
                "2nd Half Required": st.column_config.NumberColumn("2nd Half Count", min_value=0, step=1, required=True),
            },
            hide_index=True,
            use_container_width=True,
            key="requirements_editor"
        )
        
        for _, row in edited_req_df.iterrows():
            st.session_state.req_data[row["Date"]] = {
                "Session Schedule": row["Session Schedule"],
                "1st Half Required": int(row["1st Half Required"]),
                "2nd Half Required": int(row["2nd Half Required"])
            }

    st.markdown("---")
    
    # --- GL (Guest/General Lecturers) OPT-IN MANAGER ---
    st.subheader("👥 GL (Guest Lecturer) Opt-In Manager")
    st.markdown("By default, GL teachers are **excluded** from duty assignments unless explicitly opted-in below.")
    
    # Identify GL teachers using the 'Category' column containing 'GL'
    if "Category" in teachers_df.columns:
        gl_teachers = sorted(teachers_df[teachers_df["Category"].astype(str).str.contains("GL", case=False, na=False)]["Teacher/Entry"].unique().tolist())
    else:
        gl_teachers = []
    
    if "gl_opted_in" not in st.session_state:
        st.session_state.gl_opted_in = []

    selected_gls_to_opt = st.multiselect(
        "Select GL teachers to opt-in for exam duties:",
        options=gl_teachers,
        default=[t for t in st.session_state.gl_opted_in if t in gl_teachers],
        help="Only GL teachers appear here. Unselected GL teachers will be excluded from the assignment pool."
    )
    st.session_state.gl_opted_in = selected_gls_to_opt

    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🚨 Emergency Leave Manager")
        st.markdown("Select teachers and specific exam dates they are taking emergency leave.")
        
        date_strs = st.session_state.exam_dates_list
        
        if date_strs and all_teachers:
            emergency_teacher = st.selectbox("Select Teacher on Leave", options=["-- Select --"] + all_teachers, key="em_teacher")
            emergency_date = st.selectbox("Select Leave Date", options=date_strs, key="em_date")
            
            if "emergency_leaves" not in st.session_state:
                st.session_state.emergency_leaves = []
                
            if st.button("Add Emergency Leave"):
                if emergency_teacher != "-- Select --":
                    leave_entry = {"Teacher": emergency_teacher, "Date": emergency_date}
                    if leave_entry not in st.session_state.emergency_leaves:
                        st.session_state.emergency_leaves.append(leave_entry)
                        st.success(f"Added leave for {emergency_teacher} on {emergency_date}.")
            
            if st.session_state.emergency_leaves:
                st.write("**Current Emergency Leaves Applied:**")
                for idx, leave in enumerate(st.session_state.emergency_leaves):
                    c_leave_text, c_leave_btn = st.columns([3, 1])
                    c_leave_text.text(f"• {leave['Teacher']} on {leave['Date']}")
                    if c_leave_btn.button("❌ Remove", key=f"rm_leave_{idx}"):
                        st.session_state.emergency_leaves.pop(idx)
                        st.rerun()
                
                if st.button("Clear All Emergency Leaves"):
                    st.session_state.emergency_leaves = []
                    st.rerun()
        else:
            st.info("Please add at least one exam date above first.")

    with col2:
        st.subheader("⚙️ Half-Day Preferences Manager")
        st.markdown("Customize or check individual half-day preferences.")
        
        pref_init = teachers_df[["Teacher/Entry"]].drop_duplicates().reset_index(drop=True)
        if "Preferred Half" not in pref_init.columns:
            pref_init["Preferred Half"] = "No Preference"
            
        edited_prefs = st.data_editor(
            pref_init, 
            column_config={
                "Preferred Half": st.column_config.SelectboxColumn(
                    "Preferred Half",
                    options=["1st Half", "2nd Half", "No Preference"],
                    required=True
                )
            },
            hide_index=True,
            use_container_width=True,
            key="pref_editor"
        )
        preference_map = dict(zip(edited_prefs["Teacher/Entry"], edited_prefs["Preferred Half"]))

    st.markdown("---")

    # --- Helper Functions for Word & Excel Generation ---
    def set_cell_background(cell, fill_color):
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), fill_color)
        tcPr.append(shd)

    def set_table_borders(table, color="B0B0B0", sz="4", val="single"):
        tblPr = table._tbl.tblPr
        tblBorders = OxmlElement('w:tblBorders')
        for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), val)
            border.set(qn('w:sz'), sz)
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), color)
            tblBorders.append(border)
        tblPr.append(tblBorders)

    def generate_roster_docx(df):
        doc = docx.Document()
        
        for section in doc.sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1)
            section.right_margin = Inches(1)
            
        p_title = doc.add_paragraph()
        p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        min_date = pd.to_datetime(df["Date"].min()).strftime('%d/%m/%y')
        max_date = pd.to_datetime(df["Date"].max()).strftime('%d/%m/%y')
        run_title = p_title.add_run(f"INVIGILLATION DUTY CHART FROM {min_date} TO {max_date}")
        run_title.bold = True
        run_title.font.size = Pt(13)
        run_title.font.color.rgb = RGBColor(43, 88, 63)
        
        doc.add_paragraph()
        
        table = doc.add_table(rows=1, cols=4)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        set_table_borders(table)
        
        hdr_cells = table.rows[0].cells
        headers = ["DATE", "Time & Rooms", "DIC", "INVIGILATORS"]
        for i, header_text in enumerate(headers):
            hdr_cells[i].text = header_text
            set_cell_background(hdr_cells[i], '2B583F')
            p = hdr_cells[i].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            for run in p.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(10)
                
        grouped = df.groupby(["Date", "Day", "Half"])
        for (date_val, day_val, half_val), group in grouped:
            row_cells = table.add_row().cells
            dt_formatted = pd.to_datetime(date_val).strftime("%d.%m.%y")
            
            row_cells[0].text = f"{dt_formatted}\n{day_val}"
            row_cells[1].text = f"{half_val}"
            row_cells[2].text = ""
            
            invig_list = group["Invigilator Name"].tolist()
            invig_str = ", ".join(invig_list) + f" ({len(invig_list)})"
            row_cells[3].text = invig_str
            
            for i in range(4):
                p = row_cells[i].paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in p.runs:
                    run.font.size = Pt(9.5)
                    
        doc.add_paragraph()
        p_sig = doc.add_paragraph()
        p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run_sig = p_sig.add_run("Principal\nPrabhu Jagatbandhu College")
        run_sig.bold = True
        run_sig.font.size = Pt(11)
        
        buffer = io.BytesIO()
        doc.save(buffer)
        return buffer.getvalue()

    def generate_excel_summary(summary_table):
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            summary_table.to_excel(writer, index=False, sheet_name='Teacher Duty Count')
        return output.getvalue()

    # --- Roster Generation Trigger Button ---
    if st.button("🚀 Generate Duty Roster", type="primary", use_container_width=True):
        dates_list = st.session_state.get("exam_dates_list", [])
        if not dates_list:
            st.warning("Please add at least one exam date above.")
            st.session_state.roster_generated = False
        else:
            roster_records = []
            teacher_duty_counts = {name: 0 for name in all_teachers}
            generation_failed = False

            active_leaves = st.session_state.get("emergency_leaves", [])
            req_dict = st.session_state.get("req_data", {})
            opted_gls = st.session_state.get("gl_opted_in", [])

            for date_str in dates_list:
                dt_obj = pd.to_datetime(date_str)
                day_name = dt_obj.strftime("%A")
                
                day_available_rows = teachers_df[teachers_df["Day"].str.strip().str.lower() == day_name.lower()]
                day_pool = day_available_rows["Teacher/Entry"].unique().tolist()

                config = req_dict.get(date_str, {"Session Schedule": "Both Halves", "1st Half Required": 5, "2nd Half Required": 5})
                schedule_type = config.get("Session Schedule", "Both Halves")

                halves_to_run = []
                if schedule_type in ["Both Halves", "1st Half Only"]:
                    halves_to_run.append(("1st Half", config.get("1st Half Required", 5)))
                if schedule_type in ["Both Halves", "2nd Half Only"]:
                    halves_to_run.append(("2nd Half", config.get("2nd Half Required", 5)))

                for half, required_count in halves_to_run:
                    if required_count <= 0:
                        continue

                    available_teachers = []
                    for t in day_pool:
                        is_on_leave = any(l["Teacher"] == t and l["Date"] == date_str for l in active_leaves)
                        if is_on_leave:
                            continue
                        
                        is_gl = t in gl_teachers
                        if is_gl and t not in opted_gls:
                            continue

                        available_teachers.append(t)

                    if len(available_teachers) < required_count:
                        st.error(f"❌ Not enough available teachers on {day_name} ({date_str}) for {half}. Required: {required_count}, Available: {len(available_teachers)}")
                        generation_failed = True
                        break

                    def sort_key(teacher_name):
                        pref = preference_map.get(teacher_name, "No Preference")
                        pref_score = 0 if (pref == half or pref == "No Preference") else 1
                        return (pref_score, teacher_duty_counts[teacher_name])

                    available_teachers.sort(key=sort_key)
                    random.shuffle(available_teachers)
                    available_teachers.sort(key=lambda t: teacher_duty_counts[t])

                    assigned = available_teachers[:required_count]

                    for teacher in assigned:
                        teacher_duty_counts[teacher] += 1
                        roster_records.append({
                            "Date": date_str,
                            "Day": day_name,
                            "Half": half,
                            "Invigilator Name": teacher
                        })
                if generation_failed:
                    break

            if not generation_failed and roster_records:
                st.session_state.roster_df = pd.DataFrame(roster_records)
                st.session_state.summary_df = pd.DataFrame(list(teacher_duty_counts.items()), columns=["Teacher Name", "Total Duties Assigned"]).sort_values(by="Total Duties Assigned", ascending=False)
                
                st.session_state.docx_bytes = generate_roster_docx(st.session_state.roster_df)
                st.session_state.excel_bytes = generate_excel_summary(st.session_state.summary_df)
                st.session_state.roster_generated = True

    # --- Persistently Display Results & Download Buttons from Session State ---
    if st.session_state.get("roster_generated", False):
        st.success("✅ Duty Roster Generated Successfully!")
        st.subheader("📅 Comprehensive Exam Roster")
        st.dataframe(st.session_state.roster_df, use_container_width=True)

        st.subheader("📊 Overall Duty Distribution Summary")
        
        # Fixed Altair chart with clear domain bounds and integer-only tick steps
        max_duties = int(st.session_state.summary_df["Total Duties Assigned"].max())
        upper_bound = max(2, max_duties + 1)
        
        chart = alt.Chart(st.session_state.summary_df).mark_bar(color='#2b583f').encode(
            x=alt.X('Teacher Name:N', sort='-y', title='Teacher Name'),
            y=alt.Y(
                'Total Duties Assigned:Q', 
                axis=alt.Axis(
                    format='d', 
                    values=list(range(upper_bound + 1))
                ),
                scale=alt.Scale(domain=[0, upper_bound]),
                title='Total Duties Assigned'
            )
        ).properties(height=400)
        
        st.altair_chart(chart, use_container_width=True)

        col_dl1, col_dl2 = st.columns(2)
        
        with col_dl1:
            st.download_button(
                label="📥 Download Duty Chart (.docx)",
                data=st.session_state.docx_bytes,
                file_name='Duty_Chart_Roster.docx',
                mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            )
            
        with col_dl2:
            st.download_button(
                label="📊 Download Duty Counts (.xlsx)",
                data=st.session_state.excel_bytes,
                file_name='Teacher_Duty_Counts.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )
else:
    st.info("👈 Please upload your Teacher Availability CSV file using the sidebar to begin.")

# --- FREQUENTLY ASKED QUESTIONS (FAQ) SECTION ---
st.markdown("---")
st.subheader("❓ Frequently Asked Questions (FAQ)")

faq_expander1 = st.expander("1. Why are Guest Lecturers (GL) excluded by default?")
with faq_expander1:
    st.write("By institutional guidelines, Guest Lecturers are automatically filtered out from mandatory duties unless they are explicitly opted-in using the **GL Opt-In Manager** panel.")

faq_expander2 = st.expander("2. How does the app handle workload balancing?")
with faq_expander2:
    st.write("The allocation algorithm systematically tracks total duties assigned to each faculty member, prioritizing those with fewer accumulated duties while respecting half-day shift preferences.")

faq_expander3 = st.expander("3. Can I customize how many teachers are needed per shift?")
with faq_expander3:
    st.write("Yes! Once exam dates are added, a configuration table appears allowing you to modify active session schedules (*Both Halves*, *1st Half Only*, *2nd Half Only*) and change the required invigilator count per half independently.")

faq_expander4 = st.expander("4. What happens if a teacher takes emergency leave?")
with faq_expander4:
    st.write("You can log emergency leave using the **Emergency Leave Manager**. The application will instantly omit that teacher from the available pool for that specific date without altering other dates.")

# --- COPYRIGHT FOOTER ---
st.markdown("---")
st.markdown("<p style='text-align: center; color: #3a5a48; font-size: 14px;'>© Copyright reserved in favour of <b>Dr. Kisor Mukhopadhyay</b>, Prabhu Jagatbandhu College</p>", unsafe_allow_html=True)
