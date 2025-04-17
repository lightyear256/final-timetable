import pandas as pd
import random
from datetime import datetime, time, timedelta
import os

# Constants
WEEKDAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']
DAILY_BEGIN = time(9, 0)
DAILY_END = time(18, 30)
LECTURE_BLOCKS = 3  # 1.5 hours = 3 slots (30 mins each)
LAB_BLOCKS = 4      # 2 hours = 4 slots (30 mins each)
TUTORIAL_BLOCKS = 2  # 1 hour = 2 slots (30 mins each)

# Break slots configuration (30-minute slots between 12:30 to 2:00)
LUNCH_START = time(12, 30)
LUNCH_END = time(14, 0)
BREAK_DURATION = 3  # Number of 30-minute slots for lunch break
MORNING_BREAK_TIME = time(10, 30)  # Morning break remains fixed

# Department-wise break slots (each gets 1.5 hours within the lunch window)
DEPT_BREAK_SLOTS = {
    'CSE': {
        '2A': 0,  # 12:30-14:00
        '2B': 1,  # 13:00-14:30
        '4A': 0,
        '4B': 1,
        '6A': 2,  # 13:30-15:00
        '6B': 2,
        '8': 1
    },
    'DSAI': {
        '2': 1,
        '4': 2,
        '6': 0,
    },
    'ECE': {
        '2': 2,
        '4': 0,
        '6': 1,
        '8': 0
    }
}

# Color palette for different courses (pastel colors)
VISUAL_PALETTE = [
    "FFD6E0", "FFEFCF", "D6FFCF", "CFFEFF", "D6CFFF", "FFCFF4", 
    "E8D0A9", "B7E1CD", "C9DAF8", "FFD6CC", "D9D2E9", "EAD1DC",
    "A4C2F4", "D5A6BD", "B6D7A8", "FFE599", "A2C4C9", "D5D5D5"
]

# Add sections constant
CSE_SECTIONS = ['A', 'B']

# Global time periods
TIME_PERIODS = []

# Global faculty schedule tracking
faculty_schedule = {}

def initialize_time_periods():
    global TIME_PERIODS
    TIME_PERIODS = create_time_periods()

def clean_faculty_name(name):
    """Clean and standardize faculty names"""
    if pd.isna(name):
        return "TBA"
    name = str(name).strip()
    # Handle multiple faculty names (take the first one)
    if '/' in name:
        name = name.split('/')[0].strip()
    if '&' in name:
        name = name.split('&')[0].strip()
    if '(' in name:
        name = name.split('(')[0].strip()
    if 'and' in name.lower():
        name = name.split('and')[0].strip()
    return name

def initialize_faculty_schedule():
    """Initialize empty schedule for all faculty members with improved name handling"""
    global faculty_schedule
    faculty_schedule.clear()
    
    # Read faculty data
    faculty_df = pd.read_csv('faculty.csv')
    combined_df = pd.read_excel('combined2.xlsx')
    
    # Clean and standardize faculty names
    all_faculty = set(clean_faculty_name(name) for name in faculty_df['Faculty Name'].unique())
    all_faculty.update(clean_faculty_name(name) for name in combined_df['Faculty'].unique())
    
    # Remove empty or invalid names
    all_faculty = {name for name in all_faculty if name and name != "TBA"}
    
    # Initialize schedule for each faculty member
    for faculty in all_faculty:
        faculty_schedule[faculty] = {day_idx: set() for day_idx in range(len(WEEKDAYS))}

def create_course_color():
    """Generate unique colors for courses from the palette or random if needed"""
    for shade in VISUAL_PALETTE:
        yield shade
    
    # If we run out of predefined colors, generate random ones
    while True:
        red = format(random.randint(180, 255), '02x')
        green = format(random.randint(180, 255), '02x')
        blue = format(random.randint(180, 255), '02x')
        yield f"{red}{green}{blue}"

def create_time_periods():
    periods = []
    current_period = datetime.combine(datetime.today(), DAILY_BEGIN)
    finish_period = datetime.combine(datetime.today(), DAILY_END)
    
    while current_period < finish_period:
        current = current_period.time()
        next_period = current_period + timedelta(minutes=30)
        
        # Keep all time slots but we'll mark break times later
        periods.append((current, next_period.time()))
        current_period = next_period
    
    return periods

# Load data from Excel
try:
    data_frame = pd.read_excel('combined2.xlsx', sheet_name='Sheet1')
except FileNotFoundError:
    print("Error: File 'combined.xlsx' not found in the current directory")
    exit()

def get_break_slot(dept, sem, section=None):
    """Get the break slot offset for a department/semester/section"""
    if dept not in DEPT_BREAK_SLOTS:
        return 0
        
    sem_key = f"{sem}"
    if section:
        sem_key += section
    
    dept_breaks = DEPT_BREAK_SLOTS[dept]
    return dept_breaks.get(sem_key, 0)

def is_rest_period(period, dept=None, sem=None, section=None):
    """Check if a time slot falls within break times with dynamic lunch breaks"""
    begin, end = period
    
    # Morning break: 10:30-11:00
    if time(10, 30) <= begin < time(11, 0):
        return True
        
    # Dynamic lunch break
    if LUNCH_START <= begin < LUNCH_END:
        # Get department's break slot offset (0, 1, or 2 representing 30-minute offsets)
        break_offset = get_break_slot(dept, sem, section)
        break_start = datetime.combine(datetime.today(), LUNCH_START) + timedelta(minutes=30 * break_offset)
        break_end = break_start + timedelta(minutes=30 * BREAK_DURATION)
        
        # Check if current time falls within this department's break window
        current_time = datetime.combine(datetime.today(), begin)
        return break_start.time() <= begin < break_end.time()
        
    return False

def is_course_scheduled_simultaneously(schedule_grid, period, code_id):
    """Check if the same course is scheduled in any other section at the same time"""
    for day_schedule in schedule_grid.values():
        if day_schedule[period]['code'] == code_id:
            return True
    return False

def has_minimum_gap(schedule_grid, day_idx, start_period, code_id, min_gap=6):
    """Check if there's enough gap between lectures of the same course"""
    # Check backwards
    for i in range(max(0, start_period - min_gap), start_period):
        if schedule_grid[day_idx][i]['code'] == code_id:
            return False
            
    # Check forwards
    for i in range(start_period + 1, min(len(TIME_PERIODS), start_period + min_gap + 1)):
        if schedule_grid[day_idx][i]['code'] == code_id:
            return False
            
    return True

def is_faculty_available(instructor, day_idx, start_period, num_blocks):
    """Check if faculty member is available with improved name handling"""
    global faculty_schedule
    instructor = clean_faculty_name(instructor)
    
    if instructor == "TBA":
        return True
        
    if instructor not in faculty_schedule:
        faculty_schedule[instructor] = {day_idx: set() for day_idx in range(len(WEEKDAYS))}
        
    # Check if faculty is free for all required slots
    for i in range(num_blocks):
        if start_period + i in faculty_schedule[instructor][day_idx]:
            return False
    return True

def mark_faculty_busy(instructor, day_idx, start_period, num_blocks):
    """Mark faculty as busy with improved name handling"""
    global faculty_schedule
    instructor = clean_faculty_name(instructor)
    
    if instructor == "TBA":
        return
        
    for i in range(num_blocks):
        faculty_schedule[instructor][day_idx].add(start_period + i)

def is_near_break(period_idx, dept, sem, section):
    """Check if a time slot is near a break period"""
    # Check 2 slots before and after
    for i in range(max(0, period_idx - 2), min(len(TIME_PERIODS), period_idx + 3)):
        if is_rest_period(TIME_PERIODS[i], dept, sem, section):
            return True
    return False

def find_best_slot(schedule_grid, teacher_bookings, room_bookings, instructor, venue, num_blocks, day_idx, code_id, dept=None, sem=None, section=None, session_type='LEC'):
    """Find the best available time slot considering proximity to breaks"""
    global TIME_PERIODS
    best_slot = -1
    min_conflicts = float('inf')
    max_break_proximity = -1  # Track how close we are to breaks
    
    for start_period in range(len(TIME_PERIODS) - num_blocks + 1):
        conflicts = 0
        is_available = True
        break_proximity = 0  # Higher value means closer to breaks
        
        # Check basic availability
        for i in range(num_blocks):
            current_period = start_period + i
            if (current_period in teacher_bookings[instructor][day_idx] or
                current_period in room_bookings[venue][day_idx] or
                schedule_grid[day_idx][current_period]['type'] is not None or
                is_rest_period(TIME_PERIODS[current_period], dept, sem, section) or
                is_course_scheduled_simultaneously(schedule_grid, current_period, code_id) or
                not is_faculty_available(instructor, day_idx, start_period, num_blocks)):
                is_available = False
                conflicts += 1
                break
        
        # Check minimum gap requirement
        if is_available and not has_minimum_gap(schedule_grid, day_idx, start_period, code_id):
            is_available = False
            conflicts += 2
        
        # For lectures, prioritize slots near breaks
        if is_available and session_type == 'LEC':
            # Check if any slot in the block is near a break
            for i in range(num_blocks):
                if is_near_break(start_period + i, dept, sem, section):
                    break_proximity += 1
            
            # If this slot has better break proximity or same proximity but fewer conflicts
            if (break_proximity > max_break_proximity or 
                (break_proximity == max_break_proximity and conflicts < min_conflicts)):
                best_slot = start_period
                min_conflicts = conflicts
                max_break_proximity = break_proximity
        elif is_available and conflicts < min_conflicts:
            best_slot = start_period
            min_conflicts = conflicts
            
        if min_conflicts == 0 and max_break_proximity > 0:  # Found perfect slot near break
            break
    
    return best_slot

def generate_all_schedules():
    # Initialize faculty schedules at the start
    initialize_faculty_schedule()
    initialize_time_periods()
    
    # Create output directories
    output_dir = os.path.join(os.path.dirname(__file__), 'output')
    html_dir = os.path.join(output_dir, 'html')
    os.makedirs(html_dir, exist_ok=True)
    
    def is_adjacent_lecture(schedule_grid, day_idx, start_period, code_id):
        """Check if there's already a lecture of the same course in adjacent time slots"""
        # Check previous time slot
        if start_period > 0:
            prev_slot = schedule_grid[day_idx][start_period-1]
            if prev_slot['type'] == 'LEC' and prev_slot['code'] == code_id:
                return True
                
        # Check next time slot after the lecture block
        if start_period + LECTURE_BLOCKS < len(TIME_PERIODS):
            next_slot = schedule_grid[day_idx][start_period + LECTURE_BLOCKS]
            if next_slot['type'] == 'LEC' and next_slot['code'] == code_id:
                return True
                
        return False
    
    # Read both templates
    template_dir = os.path.dirname(__file__)
    with open(os.path.join(template_dir, 'template.html'), 'r', encoding='utf-8') as f:
        template = f.read()
    with open(os.path.join(template_dir, 'index_template.html'), 'r', encoding='utf-8') as f:
        index_template = f.read()
    
    # Track all generated timetables
    timetable_index = {}
    
    # Change time format
    time_format = lambda t: t.strftime("%I:%M %p")  # 12-hour format with AM/PM
    
    for dept in data_frame['Department'].unique():
        timetable_index[dept] = []
        # Get unique semesters including section info
        semesters = data_frame[data_frame['Department'] == dept]['Semester'].unique()
        
        for term in sorted(semesters, key=str):
            # Reset bookings for each semester
            teacher_bookings = {}
            room_bookings = {}
            
            # Filter subjects for this department/semester
            section_subjects = data_frame[
                (data_frame['Department'] == dept) & 
                (data_frame['Semester'] == term)
            ].copy()
            
            if section_subjects.empty:
                continue
                
            # Extract numeric semester and section if present
            term_str = str(term)
            numeric_sem = ''.join(filter(str.isdigit, term_str))
            section = term_str[-1].upper() if term_str[-1].isalpha() else None
            
            # Rest of scheduling logic
            schedule_grid = {day_idx: {period_idx: {'type': None, 'code': '', 'name': '', 'faculty': '', 'classroom': ''} 
                         for period_idx in range(len(TIME_PERIODS))} for day_idx in range(len(WEEKDAYS))}
            
            # Dictionary to store course colors
            subject_colors = {}
            color_generator = create_course_color()
            
            # First schedule all labs since they're less flexible
            practical_subjects = section_subjects[section_subjects['P'] > 0]
            for _, subject in practical_subjects.iterrows():
                code_id = str(subject['Course Code'])
                subj_name = str(subject['Course Name'])
                instructor = str(subject['Faculty'])
                regular_venue = str(subject['Classroom'])
                lab_venue = str(subject['Lab_room']) if pd.notna(subject['Lab_room']) else regular_venue
                practical_hours = int(subject['P'])
                
                # Assign a color to this course if not already assigned
                if code_id not in subject_colors:
                    subject_colors[code_id] = {"color": next(color_generator), "name": subj_name, "faculty": instructor}
                
                if instructor not in teacher_bookings:
                    teacher_bookings[instructor] = {day_idx: set() for day_idx in range(len(WEEKDAYS))}
                if lab_venue not in room_bookings:
                    room_bookings[lab_venue] = {day_idx: set() for day_idx in range(len(WEEKDAYS))}
                
                # Schedule labs - regardless of P value (2 or more), schedule only one 2-hour lab session per week
                is_scheduled = False
                try_count = 0
                while not is_scheduled and try_count < 1000:
                    day_idx = random.randint(0, len(WEEKDAYS)-1)
                    if len(TIME_PERIODS) >= LAB_BLOCKS:
                        start_period = random.randint(0, len(TIME_PERIODS)-LAB_BLOCKS)
                        
                        # Check if all required slots are free and not in break time
                        is_available = True
                        for i in range(LAB_BLOCKS):
                            if (start_period+i in teacher_bookings[instructor][day_idx] or 
                                start_period+i in room_bookings[lab_venue][day_idx] or
                                schedule_grid[day_idx][start_period+i]['type'] is not None or
                                is_rest_period(TIME_PERIODS[start_period+i])):
                                is_available = False
                                break
                        
                        if is_available:
                            # Mark professor and lab classroom as busy
                            mark_faculty_busy(instructor, day_idx, start_period, LAB_BLOCKS)
                            for i in range(LAB_BLOCKS):
                                teacher_bookings[instructor][day_idx].add(start_period+i)
                                room_bookings[lab_venue][day_idx].add(start_period+i)
                                schedule_grid[day_idx][start_period+i]['type'] = 'LAB'
                                schedule_grid[day_idx][start_period+i]['code'] = code_id if i == 0 else ''
                                schedule_grid[day_idx][start_period+i]['name'] = subj_name if i == 0 else ''
                                schedule_grid[day_idx][start_period+i]['faculty'] = instructor if i == 0 else ''
                                schedule_grid[day_idx][start_period+i]['classroom'] = lab_venue if i == 0 else ''  # Use lab venue for practical sessions
                            is_scheduled = True
                    try_count += 1
            
            # Now process all subjects that have lectures or tutorials
            theory_subjects = section_subjects[(section_subjects['L'] > 0) | (section_subjects['T'] > 0)]
            for _, subject in theory_subjects.iterrows():
                code_id = str(subject['Course Code'])
                subj_name = str(subject['Course Name'])
                instructor = str(subject['Faculty'])
                venue = str(subject['Classroom'])
                lecture_hours = int(subject['L']) if pd.notna(subject['L']) else 0
                tutorial_hours = int(subject['T']) if pd.notna(subject['T']) else 0
                
                # For 3-hour lecture courses, schedule exactly 2 lectures of 1.5 hours each
                # For 6-hour lecture courses, schedule exactly 4 lectures of 1.5 hours each
                if lecture_hours == 3:
                    num_lectures = 2
                elif lecture_hours == 6:
                    num_lectures = 4
                else:
                    num_lectures = lecture_hours
                
                # Assign a color to this course if not already assigned
                if code_id not in subject_colors:
                    subject_colors[code_id] = {"color": next(color_generator), "name": subj_name, "faculty": instructor}
                
                if instructor not in teacher_bookings:
                    teacher_bookings[instructor] = {day_idx: set() for day_idx in range(len(WEEKDAYS))}
                if venue not in room_bookings:
                    room_bookings[venue] = {day_idx: set() for day_idx in range(len(WEEKDAYS))}
                
                # Schedule lectures (1.5 hours each)
                for _ in range(num_lectures):
                    is_scheduled = False
                    try_count = 0
                    while not is_scheduled and try_count < 1000:
                        day_idx = random.randint(0, len(WEEKDAYS)-1)
                        if len(TIME_PERIODS) >= LECTURE_BLOCKS:
                            start_period = find_best_slot(
                                schedule_grid, teacher_bookings, room_bookings,
                                instructor, venue, LECTURE_BLOCKS, day_idx, code_id,
                                dept, numeric_sem, section, 'LEC'
                            )
                            
                            if start_period != -1:
                                # Mark professor and classroom as busy
                                mark_faculty_busy(instructor, day_idx, start_period, LECTURE_BLOCKS)
                                for i in range(LECTURE_BLOCKS):
                                    teacher_bookings[instructor][day_idx].add(start_period+i)
                                    room_bookings[venue][day_idx].add(start_period+i)
                                    schedule_grid[day_idx][start_period+i]['type'] = 'LEC'
                                    schedule_grid[day_idx][start_period+i]['code'] = code_id if i == 0 else ''
                                    schedule_grid[day_idx][start_period+i]['name'] = subj_name if i == 0 else ''
                                    schedule_grid[day_idx][start_period+i]['faculty'] = instructor if i == 0 else ''
                                    schedule_grid[day_idx][start_period+i]['classroom'] = venue if i == 0 else ''
                                is_scheduled = True
                        try_count += 1
                
                # Schedule tutorials (30 mins)
                for _ in range(tutorial_hours):
                    is_scheduled = False
                    try_count = 0
                    while not is_scheduled and try_count < 1000:
                        day_idx = random.randint(0, len(WEEKDAYS)-1)
                        if len(TIME_PERIODS) >= TUTORIAL_BLOCKS:
                            start_period = random.randint(0, len(TIME_PERIODS)-TUTORIAL_BLOCKS)
                            
                            # Skip if it's break time
                            if is_rest_period(TIME_PERIODS[start_period]):
                                try_count += 1
                                continue
                                
                            # Check if all required slots are free
                            is_available = True
                            for i in range(TUTORIAL_BLOCKS):
                                if (start_period+i in teacher_bookings[instructor][day_idx] or 
                                    start_period+i in room_bookings[venue][day_idx] or
                                    schedule_grid[day_idx][start_period+i]['type'] is not None):
                                    is_available = False
                                    break
                                    
                            if is_available:
                                # Mark professor and classroom as busy
                                mark_faculty_busy(instructor, day_idx, start_period, TUTORIAL_BLOCKS)
                                for i in range(TUTORIAL_BLOCKS):
                                    teacher_bookings[instructor][day_idx].add(start_period+i)
                                    room_bookings[venue][day_idx].add(start_period+i)
                                    schedule_grid[day_idx][start_period+i]['type'] = 'TUT'
                                    schedule_grid[day_idx][start_period+i]['code'] = code_id if i == 0 else ''
                                    schedule_grid[day_idx][start_period+i]['name'] = subj_name if i == 0 else ''
                                    schedule_grid[day_idx][start_period+i]['faculty'] = instructor if i == 0 else ''
                                    schedule_grid[day_idx][start_period+i]['classroom'] = venue if i == 0 else ''
                                is_scheduled = True
                        try_count += 1
            
            # Generate HTML table for this department/semester/section
            dept_content = f'''
                <div class="timetable-header">
                    <h2>{dept} Department - Semester {numeric_sem}{" - Section " + section if section else ""}</h2>
                    <p class="timestamp">Generated: {datetime.now().strftime("%d-%m-%Y %I:%M %p")}</p>
                </div>
                <table>
            '''
            
            # Add header row
            dept_content += '<tr><th>Day</th>'
            for period in TIME_PERIODS:
                dept_content += f'<th>{time_format(period[0])}<br>to<br>{time_format(period[1])}</th>'
            dept_content += '</tr>\n'
            
            # Add data rows with improved cell formatting
            for day_idx, weekday in enumerate(WEEKDAYS):
                dept_content += f'<tr><td><b>{weekday}</b></td>'
                
                skip_cells = 0
                break_count = 0
                
                # Modified rest period check to include department info
                def is_break_for_dept(period):
                    return is_rest_period(period, dept, numeric_sem, section)
                
                for period_idx in range(len(TIME_PERIODS)):
                    if skip_cells > 0:
                        skip_cells -= 1
                        continue
                    
                    # Count consecutive break periods
                    if is_break_for_dept(TIME_PERIODS[period_idx]):
                        break_count = 1
                        next_idx = period_idx + 1
                        while next_idx < len(TIME_PERIODS) and is_break_for_dept(TIME_PERIODS[next_idx]):
                            break_count += 1
                            next_idx += 1
                        dept_content += f'<td colspan="{break_count}" class="break">BREAK</td>'
                        skip_cells = break_count - 1
                    elif schedule_grid[day_idx][period_idx]['type']:
                        if schedule_grid[day_idx][period_idx]['code']:
                            session_type = schedule_grid[day_idx][period_idx]['type']
                            code_id = schedule_grid[day_idx][period_idx]['code']
                            venue = schedule_grid[day_idx][period_idx]['classroom']
                            faculty = schedule_grid[day_idx][period_idx]['faculty']
                            color = subject_colors.get(code_id, {}).get('color', 'ffffff')
                            
                            if session_type == 'LEC':
                                colspan = LECTURE_BLOCKS
                            elif session_type == 'LAB':
                                colspan = LAB_BLOCKS
                            else:  # TUT
                                colspan = TUTORIAL_BLOCKS
                                
                            skip_cells = colspan - 1
                            dept_content += f'''<td colspan="{colspan}" class="timetable-cell">
                                <div class="course-block" style="background-color: #{color}">
                                    <strong>{code_id} {session_type}</strong><br>
                                    Room: {venue}<br>
                                    {faculty}
                                </div>
                            </td>'''
                        else:
                            dept_content += '<td></td>'
                    else:
                        dept_content += '<td></td>'
                        
                dept_content += '</tr>\n'
            
            dept_content += '</table>\n'
            
            # Add LTPSC Legend first
            dept_content += '<div class="legend"><h3>LTPSC Information</h3>\n<table>\n'
            dept_content += '<tr><th>Course Code</th><th>L</th><th>T</th><th>P</th><th>S</th><th>C</th></tr>\n'
            
            for code_id in subject_colors.keys():
                course_info = section_subjects[section_subjects['Course Code'] == code_id].iloc[0]
                l_hours = int(course_info['L']) if pd.notna(course_info['L']) else 0
                t_hours = int(course_info['T']) if pd.notna(course_info['T']) else 0
                p_hours = int(course_info['P']) if pd.notna(course_info['P']) else 0
                s_hours = int(course_info['S']) if pd.notna(course_info['S']) else 0
                credits = int(course_info['C']) if pd.notna(course_info['C']) else 0
                
                dept_content += f'''
                <tr>
                    <td><strong>{code_id}</strong></td>
                    <td>{l_hours}</td>
                    <td>{t_hours}</td>
                    <td>{p_hours}</td>
                    <td>{s_hours}</td>
                    <td>{credits}</td>
                </tr>'''
            
            dept_content += '</table></div>\n'
            
            # Add Course Legend after LTPSC
            dept_content += '<div class="legend"><h3>Course Legend</h3>\n<table>\n'
            dept_content += '<tr><th>Course Code</th><th>Color</th><th>Course Name</th><th>Faculty</th></tr>\n'
            
            for code_id, details in subject_colors.items():
                dept_content += f'''
                <tr>
                    <td><strong>{code_id}</strong></td>
                    <td><div class="legend-color" style="background-color: #{details['color']}"></div></td>
                    <td>{details['name']}</td>
                    <td>{details['faculty']}</td>
                </tr>'''
            
            dept_content += '</table></div>\n'
        
            # Save semester-specific HTML file with updated path
            safe_dept = dept.replace(' ', '_').lower()
            filename = f'timetable_{safe_dept}_semester_{numeric_sem}{"_section_" + section.lower() if section else ""}.html'
            filepath = os.path.join(html_dir, filename)
            timetable_index[dept].append((numeric_sem, section, filename))  # Keep original filename for links
            
            sem_html = template.replace('<!-- CONTENT_PLACEHOLDER -->', dept_content)
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write(sem_html)
            
            print(f"Timetable for {dept} - Semester {numeric_sem}{' - Section ' + section if section else ''} has been saved to {filepath}")
            
            # Reset dept_content for next section
            dept_content = ''
    
    # Generate index page
    index_content = ''
    for dept, semesters in sorted(timetable_index.items()):
        index_content += f'''
        <div class="dept-section">
            <h2 class="dept-title">{dept} Department</h2>
            <div class="semester-grid">
        '''
        
        if dept == 'CSE':
            # Group CSE by semester
            sem_groups = {}
            for term, section, filename in sorted(semesters):
                if term not in sem_groups:
                    sem_groups[term] = []
                sem_groups[term].append((section, filename))
            
            for term in sorted(sem_groups.keys()):
                index_content += f'<div class="semester-group"><h3>Semester {term}</h3>'
                for section, filename in sorted(sem_groups[term]):
                    index_content += f'''
                        <a href="{filename}" class="semester-link">
                            Section {section}
                        </a>
                    '''
                index_content += '</div>'
        else:
            # Unchanged display for other departments
            for term, _, filename in sorted(semesters):
                index_content += f'''
                    <a href="{filename}" class="semester-link">
                        Semester {term}
                    </a>
                '''
        
        index_content += '</div></div>'

    # Save index file in html directory
    index_html = index_template.replace('<!-- CONTENT_PLACEHOLDER -->', index_content)
    index_path = os.path.join(html_dir, 'index.html')
    with open(index_path, 'w', encoding='utf-8') as f:
        f.write(index_html)
    
    print(f"Main index page generated as {index_path}")

if __name__ == "__main__":
    generate_all_schedules()