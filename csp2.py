"""
timetable_full_solver_v3.py
Reads data from a single Excel file (csit_data.xlsx), handles critical constraints 
(instructor, room, and student clashes), and incorporates soft constraints like 
instructor qualifications and preferences.
Outputs a valid timetable solution and a separate log of any lectures that could not be scheduled.
"""

import pandas as pd
import random
from collections import defaultdict
import sqlite3 # Keep for potential type hints, though not strictly needed

# -------------------------
# Load Excel file (FIXED)
# -------------------------
def load_tables_excel():
    """Load timetable data from a single Excel file."""
    
    excel_filename = "csit_data.xlsx"
    sheet_names = ["Courses", "Instructors", "Rooms", "TimeSlots", "Sections"]
    
    try:
        # Load all sheets from the Excel file
        all_sheets = pd.read_excel(excel_filename, sheet_name=sheet_names)
        
        courses_df = all_sheets["Courses"]
        instructors_df = all_sheets["Instructors"]
        rooms_df = all_sheets["Rooms"]
        timeslots_df = all_sheets["TimeSlots"]
        sections_df = all_sheets["Sections"]
        
    except Exception as e:
        print(f"⚠️ Could not read Excel file '{excel_filename}': {e}")
        print("Please ensure the file is in the same folder and the sheet names are correct:")
        print("Courses, Instructors, Rooms, TimeSlots, Sections")
        # Return empty dataframes to prevent a crash
        empty_df = pd.DataFrame()
        return empty_df, empty_df, empty_df, empty_df, empty_df, empty_df

    # We no longer build curriculum here, it will be handled in preprocess.
    # We return an empty dataframe for it to maintain the tuple structure for unpacking.
    curriculum_df = pd.DataFrame() 
    
    print("✅ Loaded all data from Excel file successfully.")
    return courses_df, instructors_df, rooms_df, timeslots_df, sections_df, curriculum_df

# -------------------------
# Helpers
# -------------------------
def safe_str(x):
    return "" if pd.isna(x) else str(x).strip()

def int_safe(x, default=0):
    try: return int(x)
    except: return default

# *** THIS IS THE FIXED FUNCTION ***
def compatible_room(course_type, room_type):
    c, r = (course_type or "").lower(), (room_type or "").lower()

    # A 'lab' course can go in a 'lab' room
    if "lab" in c and "lab" in r:
        return True
        
    # A 'lecture' course can go in a 'classroom', 'hall', or 'theater'
    if "lec" in c and ("classroom" in r or "hall" in r or "theater" in r):
        return True

    # --- THE FIX ---
    # Allow 'lab' courses (like CSC317) to also be in large rooms
    if "lab" in c and ("classroom" in r or "hall" in r or "theater" in r):
        return True

    # Fallback for any other combo
    if "classroom" in r or "hall" in r:
        return True

    return False
# *** END OF FIXED FUNCTION ***

# -------------------------
# Preprocess (FIXED column names)
# -------------------------
def preprocess(courses_df, instructors_df, rooms_df, timeslots_df, sections_df, curriculum_df):
    courses = {}
    for _, r in courses_df.iterrows():
        cid = safe_str(r.get("CourseID"))
        ctype = safe_str(r.get("Type", "")).lower()
        if not cid: continue
        
        # *** BUG FIX HERE ***
        # If a course (like CNC311) is already in the dict as a "lecture",
        # do NOT overwrite it with the "lab" entry that comes after it.
        if cid in courses and courses[cid]["type"] == "lecture":
            pass # Keep the "lecture" type
        else:
            courses[cid] = {"name": safe_str(r.get("CourseName")), "type": ctype}
        # *** END OF FIX ***

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
        rooms[rid] = {
            "type": safe_str(r.get("Type", "")).lower(),
            "capacity": int_safe(r.get("Capacity", 0))
        }

    timeslots, timeslot_info = [], {}
    for _, r in timeslots_df.iterrows():
        tid = safe_str(r.get("TimeSlotID"))
        day = safe_str(r.get("Day"))
        start = safe_str(r.get("StartTime"))
        end = safe_str(r.get("EndTime"))
        if not tid: continue
        timeslots.append(tid)
        timeslot_info[tid] = {"day": day, "start": start, "end": end}

    # *** GROUPING LOGIC ***
    print("✅ Grouping sections into lecture groups...")
    lecture_groups = []
    
    group_column = 'Group' if 'Group' in sections_df.columns else 'SectionID'
    
    grouped = sections_df.groupby(['Level', group_column])
    
    for (level, group_id), group_data in grouped:
        total_students = group_data['StudentCount'].sum()
        courses_str = safe_str(group_data['Courses'].iloc[0])
        course_list = [c.strip() for c in courses_str.split(',') if c.strip()]
        
        lecture_groups.append({
            "group_name": f"{level}_{group_id}", # e.g., "1_1" or "3_AI_1"
            "year": level,
            "students": total_students,
            "courses": course_list
        })
        
    print(f"✅ Combined {len(sections_df)} sections into {len(lecture_groups)} lecture groups.")
    return courses, instructors, rooms, timeslots, timeslot_info, lecture_groups

# -------------------------
# LectureVar class
# -------------------------
class LectureVar:
    def __init__(self, course, section, year, students):
        self.course = course
        self.section = section
        self.year = year
        self.students = students
        self.name = f"{course}_{section}"

    def __repr__(self):
        return self.name

# -------------------------
# Build variables & domains
# -------------------------
def build_vars_domains(courses, instructors, rooms, timeslots, lecture_groups):
    variables, domains = [], {}
    
    # *** GROUPING LOGIC FIX ***
    # Iterate over 'lecture_groups' instead of 'sections'
    for group in lecture_groups:
        year, students, group_name = group["year"], group["students"], group["group_name"]
        
        # Get courses specific to THIS group
        for cid in group.get("courses", []):
            course_info = courses.get(cid)
            if not course_info:
                print(f"⚠️ Warning: Course {cid} for Group {group_name} not in Courses sheet. Skipping.")
                continue

            ctype = course_info.get("type", "")
            
            # Create one LectureVar per group, not per section
            v = LectureVar(cid, group_name, year, students)
            variables.append(v)
            dom = []
            
            for t in timeslots:
                for r_id, r_info in rooms.items():
                    if not compatible_room(ctype, r_info["type"]): 
                        continue
                    # Check capacity against the SUM of students in the group
                    if r_info["capacity"] < students: 
                        continue
                        
                    for instr_id, i_info in instructors.items():
                        is_qualified = cid in i_info["quals"]
                        is_preferred = t in i_info["prefs"]
                        
                        if not is_qualified:
                           continue
                        
                        dom.append((t, r_id, instr_id, is_qualified, is_preferred))
                        
            domains[v] = dom
    return variables, domains

# -------------------------
# Greedy Backtracking Solver
# -------------------------
def solve_timetable(variables, domains):
    assigned = {}
    failed_assignments = []

    # Sets to track resource usage: (timeslot, resource_id)
    used_room_ts = set()
    used_instr_ts = set()
    used_section_ts = set() # NEW: Prevents student clashes

    # Heuristic: Schedule the most constrained lectures first (e.g., largest classes)
    sorted_vars = sorted(variables, key=lambda v: (-v.students, len(domains.get(v, []))))

    for v in sorted_vars:
        dom = domains.get(v, [])
        if not dom:
            print(f"🔴 FAILED to schedule: {v.name} (Students: {v.students}) - NO DOMAIN (No qualified instructor/room)")
            failed_assignments.append(v)
            continue
            
        # Heuristic: Try best options first (qualified & preferred)
        sorted_domain = sorted(dom, key=lambda x: (x[3], x[4]), reverse=True)
        
        assigned_slot = False
        for option in sorted_domain:
            t, r, instr, _, _ = option
            
            # CHECK ALL HARD CONSTRAINTS
            if (t, r) in used_room_ts: continue
            if (t, instr) in used_instr_ts: continue
            if (t, v.section) in used_section_ts: continue # CRITICAL: Student clash check

            # If all checks pass, assign it
            assigned[v] = option
            used_room_ts.add((t, r))
            used_instr_ts.add((t, instr))
            used_section_ts.add((t, v.section))
            assigned_slot = True
            break
        
        # If no valid slot was found after checking all options
        if not assigned_slot:
            failed_assignments.append(v)
            print(f"🔴 FAILED to schedule: {v.name} (Students: {v.students}) - All slots clashed")

    return assigned, failed_assignments

# -------------------------
# Export CSV
# -------------------------
def export_results(assigned, failed, timeslot_info, instructors):
    # --- Successful Assignments ---
    rows = []
    for v, val in assigned.items():
        if not val: continue
        t, r, instr_id, qual, pref = val
        info = timeslot_info.get(t, {"day": "N/A", "start": "N/A", "end": "N/A"})
        instr_name = instructors.get(instr_id, {}).get("name", "N/A")
        rows.append({
            "Course": v.course,
            "Section": v.section,
            "Year": v.year,
            "Students": v.students,
            "Day": info["day"],
            "Start": info["start"],
            "End": info["end"],
            "Room": r,
            "Instructor": instr_name,
            "InstructorQualified": bool(qual),
            "TimeslotIsPreferred": bool(pref)
        })
    
    solution_file = "timetable_solution.csv"
    pd.DataFrame(rows).to_csv(solution_file, index=False)
    print(f"✅ Exported {len(rows)} successful assignments to {solution_file}")

    # --- Failed Assignments ---
    if failed:
        failed_rows = [{"Course": v.course, "Section": v.section, "Year": v.year, "Students": v.students} for v in failed]
        failures_file = "timetable_failures.csv"
        pd.DataFrame(failed_rows).to_csv(failures_file, index=False)
        print(f"⚠️ Exported {len(failed_rows)} failed assignments to {failures_file}")

# -------------------------
# Main Execution (FIXED)
# -------------------------
def main():
    print("📘 Loading data from Excel file ...")
    all_data = load_tables_excel()
    
    if all_data[0].empty:
        print("❌ Data loading failed. Exiting.")
        return

    print("⚙️ Preprocessing data ...")
    # *** GROUPING LOGIC FIX *** (Return 'lecture_groups' instead of 'sections')
    courses, instructors, rooms, timeslots, t_info, lecture_groups = preprocess(*all_data)
    print(f"📊 Data ready: {len(courses)} courses, {len(instructors)} instructors, {len(rooms)} rooms, {len(timeslots)} timeslots.")

    print("🧩 Building variables and domains ...")
    # *** GROUPING LOGIC FIX *** (Pass 'lecture_groups')
    variables, domains = build_vars_domains(courses, instructors, rooms, timeslots, lecture_groups)
    print(f"✅ Created {len(variables)} lecture variables to schedule.")
    
    if not variables:
        print("❌ No lecture variables were created. Check your 'Sections' and 'Courses' data.")
        print("🎉 Done.")
        return

    print("🧠 Solving timetable... please wait.")
    assigned, failed = solve_timetable(variables, domains)
    
    print("📄 Exporting results...")
    export_results(assigned, failed, t_info, instructors)
    
    print("🎉 Done.")

if __name__ == "__main__":
    main()
