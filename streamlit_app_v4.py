import streamlit as st
import pandas as pd
from collections import defaultdict
import time
import random

# ----------------------------------------------------------------------
# 1. ALL THE SOLVER LOGIC
# ----------------------------------------------------------------------

# -------------------------
# Helpers
# -------------------------
def safe_str(x):
    return "" if pd.isna(x) else str(x).strip()

def int_safe(x, default=0):
    try: return int(x)
    except: return default

def compatible_room(course_type, room_type):
    c, r = (course_type or "").lower(), (room_type or "").lower()
    if "lab" in c and "lab" in r: return True
    if "lec" in c and ("classroom" in r or "hall" in r or "theater" in r): return True
    if "lab" in c and ("classroom" in r or "hall" in r or "theater" in r): return True
    if "classroom" in r or "hall" in r: return True
    return False

# -------------------------
# LectureVar class (***UPDATED***)
# -------------------------
class LectureVar:
    def __init__(self, course, group_name, year, students, section_display_name):
        self.course = course
        self.group_name = group_name  # Internal ID, e.g., "1_1"
        self.year = year
        self.students = students
        self.section_display_name = section_display_name # Display string, e.g., "(Sec 1, 2, 3)"
        self.name = f"{course}_{group_name}"
        
    def __repr__(self): return self.name

# -------------------------
# Load Excel (Modified for Streamlit)
# -------------------------
def load_tables_excel(uploaded_file):
    sheet_names = ["Courses", "Instructors", "Rooms", "TimeSlots", "Sections"]
    try:
        all_sheets = pd.read_excel(uploaded_file, sheet_name=sheet_names)
        return (
            all_sheets["Courses"], all_sheets["Instructors"], all_sheets["Rooms"],
            all_sheets["TimeSlots"], all_sheets["Sections"], pd.DataFrame()
        )
    except Exception as e:
        st.error(f"⚠️ Could not read Excel file: {e}")
        st.error("Please ensure the file has the correct sheets: Courses, Instructors, Rooms, TimeSlots, Sections")
        return (pd.DataFrame(),) * 6

# -------------------------
# Preprocess (***UPDATED***)
# -------------------------
def preprocess(courses_df, instructors_df, rooms_df, timeslots_df, sections_df, curriculum_df):
    courses = {}
    for _, r in courses_df.iterrows():
        cid = safe_str(r.get("CourseID"))
        ctype = safe_str(r.get("Type", "")).lower()
        if not cid: continue
        if cid in courses and courses[cid]["type"] == "lecture": pass
        else: courses[cid] = {"name": safe_str(r.get("CourseName")), "type": ctype}

    instructors = {}
    for _, r in instructors_df.iterrows():
        iid = safe_str(r.get("InstructorID"))
        if not iid: continue
        quals_raw = safe_str(r.get("QualifiedCourses", "")) 
        prefs_raw = safe_str(r.get("preferred_slots", ""))
        instructors[iid] = {
            "name": safe_str(r.get("Name")),
            "quals": set(q.strip() for q in quals_raw.split(',') if q.strip()),
            "prefs": set(p.strip() for p in prefs_raw.split(',') if p.strip())
        }

    rooms = {}
    for _, r in rooms_df.iterrows():
        rid = safe_str(r.get("RoomID"))
        if not rid: continue
        rooms[rid] = {"type": safe_str(r.get("Type", "")).lower(), "capacity": int_safe(r.get("Capacity", 0))}

    timeslots, timeslot_info = [], {}
    for _, r in timeslots_df.iterrows():
        tid = safe_str(r.get("TimeSlotID"))
        day = safe_str(r.get("Day"))
        start = safe_str(r.get("StartTime"))
        end = safe_str(r.get("EndTime"))
        if not tid: continue
        timeslots.append(tid)
        timeslot_info[tid] = {"day": day, "start": str(start), "end": str(end)}

    # --- NEW: Grouping logic to get section display names ---
    lecture_groups = []
    group_column = 'Group' if 'Group' in sections_df.columns else 'SectionID'
    grouped = sections_df.groupby(['Level', group_column])
    
    for (level, group_id), group_data in grouped:
        total_students = group_data['StudentCount'].sum()
        courses_str = safe_str(group_data['Courses'].iloc[0])
        course_list = [c.strip() for c in courses_str.split(',') if c.strip()]
        
        # --- Create the new display name ---
        sec_ids = group_data['SectionID'].astype(str).tolist()
        section_display_name = f"(Sec {', '.join(sec_ids)})"
        
        lecture_groups.append({
            "group_name": f"{level}_{group_id}", # Internal ID
            "section_display_name": section_display_name, # Display Name
            "year": level,
            "students": total_students,
            "courses": course_list
        })
        
    return courses, instructors, rooms, timeslots, timeslot_info, lecture_groups

# -------------------------
# Build variables & domains (***UPDATED***)
# -------------------------
def build_vars_domains(courses, instructors, rooms, timeslots, lecture_groups):
    variables, domains = [], {}
    for group in lecture_groups:
        year, students, group_name = group["year"], group["students"], group["group_name"]
        section_display_name = group["section_display_name"]
        
        for cid in group.get("courses", []):
            course_info = courses.get(cid)
            if not course_info:
                st.warning(f"⚠️ Warning: Course {cid} for Group {group_name} not in Courses sheet. Skipping.")
                continue
            ctype = course_info.get("type", "")
            
            # --- Pass new display name to LectureVar ---
            v = LectureVar(cid, group_name, year, students, section_display_name)
            variables.append(v)
            
            dom = []
            for t in timeslots:
                for r_id, r_info in rooms.items():
                    if not compatible_room(ctype, r_info["type"]): continue
                    if r_info["capacity"] < students: continue
                    for instr_id, i_info in instructors.items():
                        is_qualified = cid in i_info["quals"]
                        if not is_qualified: continue
                        is_preferred = t in i_info["prefs"]
                        dom.append((t, r_id, instr_id, is_qualified, is_preferred))
            
            random.shuffle(dom)
            domains[v] = dom
    return variables, domains

# -------------------------
# Greedy Backtracking Solver
# -------------------------
def solve_timetable(variables, domains):
    assigned, failed_assignments = {}, []
    used_room_ts, used_instr_ts, used_section_ts = set(), set(), set()
    sorted_vars = sorted(variables, key=lambda v: (-v.students, len(domains.get(v, []))))

    for v in sorted_vars:
        dom = domains.get(v, [])
        if not dom:
            st.warning(f"🔴 FAILED: {v.name} (Students: {v.students}) - NO DOMAIN")
            failed_assignments.append(v); continue
        
        assigned_slot = False
        for option in dom: # Use the shuffled domain
            t, r, instr, _, _ = option
            # --- Use internal group_name for clash check ---
            if (t, r) in used_room_ts: continue
            if (t, instr) in used_instr_ts: continue
            if (t, v.group_name) in used_section_ts: continue
            
            assigned[v] = option
            used_room_ts.add((t, r)); used_instr_ts.add((t, instr)); used_section_ts.add((t, v.group_name))
            assigned_slot = True; break
        if not assigned_slot:
            st.warning(f"🔴 FAILED: {v.name} (Students: {v.students}) - All slots clashed")
            failed_assignments.append(v)
    return assigned, failed_assignments

# -------------------------
# Format Results (***UPDATED***)
# -------------------------
def format_results(assigned, failed, timeslot_info, instructors, courses):
    rows = []
    for v, val in assigned.items():
        if not val: continue
        t, r, instr_id, qual, pref = val
        info = timeslot_info.get(t, {"day": "N/A", "start": "N/A", "end": "N/A"})
        instr_name = instructors.get(instr_id, {}).get("name", "N/A")
        ctype = courses.get(v.course, {}).get("type", "unknown")
        
        rows.append({
            "Course": v.course,
            "Group": v.section_display_name, # <-- Use the new display name
            "Year": v.year, "Students": v.students,
            "Day": info["day"], "Start": info["start"], "End": info["end"], "Room": r,
            "Instructor": instr_name, "InstructorQualified": bool(qual), "TimeslotIsPreferred": bool(pref),
            "Type": ctype
        })
    solution_df = pd.DataFrame(rows)
    
    # --- Format failed assignments ---
    failed_rows = [{
        "Course": v.course, 
        "Group": v.section_display_name, # <-- Use the new display name
        "Year": v.year, 
        "Students": v.students
    } for v in failed]
    failures_df = pd.DataFrame(failed_rows)
    
    return solution_df, failures_df

# ----------------------------------------------------------------------
# 2. FUNCTION TO CREATE GRID TABLES (***UPDATED***)
# ----------------------------------------------------------------------
def format_as_grid_tables(solution_df, timeslot_info):
    if solution_df.empty:
        return {}

    # --- Helper function to create HTML for each cell ---
    def create_display_html(row):
        # Set color class based on Type
        course_type = "lab" if "lab" in row['Type'] else "lec"
        return (
            f"<div class='cell-content type-{course_type}'>"
            f"<strong>{row['Course']}</strong>"
            f"<em class='group-name'>{row['Group']}</em>"
            f"<span class='room'>@ {row['Room']}</span>"
            f"<span class='instructor-name'>{row['Instructor']}</span>"
            "</div>"
        )
    # --- END HELPER ---

    solution_df['TimeSlot'] = solution_df['Start'].astype(str) + '-' + solution_df['End'].astype(str)
    solution_df['Display'] = solution_df.apply(create_display_html, axis=1)

    day_order = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday"]
    
    all_time_slots = []
    seen_starts = set()
    sorted_times = sorted(timeslot_info.values(), key=lambda x: (x['start']))
    for info in sorted_times:
        if info['start'] not in seen_starts:
            all_time_slots.append(f"{info['start']}-{info['end']}")
            seen_starts.add(info['start'])
    time_order = all_time_slots

    tables_by_year = {}
    all_years = sorted(solution_df['Year'].unique())

    for year in all_years:
        year_df = solution_df[solution_df['Year'] == year]
        
        try:
            pivot_table = year_df.pivot_table(
                index='TimeSlot', columns='Day', values='Display',
                aggfunc=lambda x: "".join(x) # Join without <br>, let flexbox handle it
            )
        except Exception:
            continue
        
        pivot_table = pivot_table.reindex(index=time_order, columns=day_order)
        pivot_table.index.name = "Time"
        pivot_table = pivot_table.fillna("")
        
        html_table = pivot_table.to_html(escape=False, border=0, classes=["timetable-grid"])
        tables_by_year[year] = html_table
            
    return tables_by_year

# ----------------------------------------------------------------------
# 3. THE STREAMLIT APP GUI (***BEAUTIFUL CSS***)
# ----------------------------------------------------------------------

st.set_page_config(page_title="Timetable Solver", layout="wide")

# --- CUSTOM CSS TO MAKE THE TABLE READABLE ---
st.markdown("""
<style>
/* --- Main App Body --- */
body {
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
}
/* --- Main Table Container --- */
.timetable-grid {
    border-collapse: collapse; /* Use a classic grid */
    width: 100%;
    font-size: 13px;
    border-radius: 8px; /* Rounded corners for the whole table */
    overflow: hidden; /* Clips the inner content to the border-radius */
    border: 1px solid #e2e8f0; /* Light gray border for the outside */
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
}
/* --- All Cells: Headers and Data --- */
.timetable-grid th, .timetable-grid td {
    border: 1px solid #e2e8f0; /* Thin, light gray border for all slots */
    text-align: left;
    vertical-align: top;
    height: auto; 
    width: 16.66%;
    min-height: 140px;
    padding: 0; 
}

/* --- Table headers (Days) and Time Slot label --- */
.timetable-grid th {
    background-color: #ffffff; /* --- REMOVED THE "BLOCK" --- */
    color: #475569; /* Dark Gray-Blue Text */
    font-size: 14px;
    font-weight: 600;
    padding: 12px 8px;
    text-align: center;
    vertical-align: middle;
}
/* Table index (Times) */
.timetable-grid .index {
    font-weight: 600;
    font-size: 14px;
    background-color: #ffffff; /* --- REMOVED THE "BLOCK" --- */
    color: #475569;
    text-align: center;
    vertical-align: middle;
    padding: 8px;
}

/* Top-left corner cell */
.timetable-grid th.index_name {
    background-color: #ffffff; /* --- REMOVED THE "BLOCK" --- */
    color: #475569;
}
.timetable-grid th.index_name::before {
    content: "Time Slot";
    font-size: 14px;
    font-weight: 600;
}

/* --- Style for "white squares" (empty cells) --- */
.timetable-grid td:empty {
    background-color: #ffffff; /* Plain white */
    min-height: 20px;
    height: auto;
}

/* --- Style for cells WITH content --- */
.cell-content {
    padding: 10px;
    height: 100%;
    min-height: 140px;
    display: flex;
    flex-direction: column;
    justify-content: flex-start;
    gap: 6px; /* Space between elements */
    background-color: #ffffff; /* White background for the card */
}
/* Course Code */
.cell-content strong {
    font-size: 14px;
    color: #111827; /* Darker text for course */
    font-weight: 700;
}
/* Group Name (e.g., Sec 1, 2, 3) */
.group-name {
    font-size: 13px;
    color: #4b5563; /* Gray text */
    font-style: italic;
    font-weight: 500;
}
/* Room */
.room {
    font-size: 13px;
    color: #4b5563; /* Gray text */
}
/* Instructor Name */
.instructor-name {
    font-size: 13px;
    font-weight: 500;
    color: #1d4ed8; /* Blue text for instructor */
}

/* --- NEW: Color coding for Lec/Lab --- */
.type-lec {
    background-color: #fefce8; /* Softer Light Yellow */
    border-left: 4px solid #f59e0b; /* Yellow Accent */
}
.type-lab {
    background-color: #f0fdf4; /* Softer Light Green */
    border-left: 4px solid #22c55e; /* Green Accent */
}

/* --- Remove the black lines --- */
hr { display: none; }
</style>
""", unsafe_allow_html=True)
# --- END OF CSS ---

st.title("🎓 University Timetable Solver")
st.markdown("This app generates a valid timetable based on your university's data (courses, instructors, rooms, etc.) using a Constraint Satisfaction Problem (CSP) solver.")

# --- Sidebar for Upload ---
st.sidebar.title("Controls")
uploaded_file = st.sidebar.file_uploader("Upload Your Excel File", type=["xlsx"])

if uploaded_file is None:
    st.info("👈 Please upload your `csit_data.xlsx` file in the sidebar to begin.")
    st.image("https://placehold.co/800x400/e2e8f0/64748b?text=Upload+Your+Data+File", caption="Upload your Excel data file to generate the timetable.")

if uploaded_file:
    st.sidebar.success(f"Uploaded: `{uploaded_file.name}`")

    if st.sidebar.button("Generate Timetable", type="primary", use_container_width=True):
        
        start_time = time.time()
        
        with st.spinner("Loading and processing data..."):
            all_data = load_tables_excel(uploaded_file)
            if all_data[0].empty: st.stop()
            courses, instructors, rooms, timeslots, t_info, lecture_groups = preprocess(*all_data)
            st.text(f"📊 Data ready: {len(courses)} courses, {len(instructors)} instructors, {len(rooms)} rooms.")
            variables, domains = build_vars_domains(courses, instructors, rooms, timeslots, lecture_groups)
            st.text(f"✅ Created {len(variables)} lecture groups to schedule.")

        with st.spinner("🧠 Solving timetable... this may take a moment."):
            assigned, failed = solve_timetable(variables, domains)
            solution_df, failures_df = format_results(assigned, failed, t_info, instructors, courses)
        
        end_time = time.time()
        st.success("🎉 Timetable generation complete!")
        
        # --- Performance Metrics ---
        st.subheader("Performance Metrics")
        col1, col2, col3 = st.columns(3)
        col1.metric("Generation Time", f"{end_time - start_time:.2f} s")
        col2.metric("Successful Assignments", f"{len(solution_df)}")
        col3.metric("Failed Assignments", f"{len(failures_df)}")
        
        # --- Display HTML Grid Timetables ---
        st.header("Visual Timetable Grids")
        st.markdown("This is the final timetable, formatted as a grid for each level.")
        
        grid_tables = format_as_grid_tables(solution_df, t_info) 
        
        if not grid_tables:
            st.warning("No successful assignments to display in a grid.")
        else:
            tab_titles = [f"Level {year}" for year in grid_tables.keys()]
            tabs = st.tabs(tab_titles)
            
            for i, (year, table_html) in enumerate(grid_tables.items()):
                with tabs[i]:
                    st.markdown(table_html, unsafe_allow_html=True)
        
        st.divider()

        # --- Failed Assignments ---
        if not failures_df.empty:
            st.header(f"⚠️ Failed Assignments ({len(failures_df)})")
            st.dataframe(failures_df, use_container_width=True)
        else:
            st.header("🎉 All lectures scheduled successfully! (0 failures)")

        # --- Raw Data (in expander) ---
        with st.expander(f"View Raw Solution Data ({len(solution_df)} rows)"):
            st.dataframe(solution_df, use_container_width=True)
            st.download_button(
                label="Download Raw Solution as CSV",
                data=solution_df.to_csv(index=False).encode('utf-8'),
                file_name='timetable_solution.csv',
                mime='text/csv',
            )

