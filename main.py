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

# Color palette for different courses (pastel colors)
VISUAL_PALETTE = [
    "FFD6E0", "FFEFCF", "D6FFCF", "CFFEFF", "D6CFFF", "FFCFF4", 
    "E8D0A9", "B7E1CD", "C9DAF8", "FFD6CC", "D9D2E9", "EAD1DC",
    "A4C2F4", "D5A6BD", "B6D7A8", "FFE599", "A2C4C9", "D5D5D5"
]

# Add sections constant
CSE_SECTIONS = ['A', 'B']

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
    data_frame = pd.read_excel('combined.xlsx', sheet_name='Sheet1')
except FileNotFoundError:
    print("Error: File 'combined.xlsx' not found in the current directory")
    exit()

def is_rest_period(period):
    """Check if a time slot falls within break times"""
    begin, end = period
    # Morning break: 10:30-11:00
    morning_rest = (time(10, 30) <= begin < time(11, 0))
    # Lunch break: 12:30-14:30
    lunch_rest = (time(12, 30) <= begin < time(14, 30))
    return morning_rest or lunch_rest

def generate_all_schedules():
    # Create output directories
    output_dir = os.path.join(os.path.dirname(__file__), 'output')
    html_dir = os.path.join(output_dir, 'html')
    os.makedirs(html_dir, exist_ok=True)
    
    TIME_PERIODS = create_time_periods()
    
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
            
            # Process section-specific subjects
            practical_subjects = section_subjects[section_subjects['P'] > 0]
            for _, subject in practical_subjects.iterrows():
                code_id = str(subject['Course Code'])
                subj_name = str(subject['Course Name'])
                instructor = str(subject['Faculty'])
                venue = str(subject['Classroom'])
                practical_hours = int(subject['P'])
                
                # Assign a color to this course if not already assigned
                if code_id not in subject_colors:
                    subject_colors[code_id] = {"color": next(color_generator), "name": subj_name, "faculty": instructor}
                
                if instructor not in teacher_bookings:
                    teacher_bookings[instructor] = {day_idx: set() for day_idx in range(len(WEEKDAYS))}
                if venue not in room_bookings:
                    room_bookings[venue] = {day_idx: set() for day_idx in range(len(WEEKDAYS))}
                
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
                                start_period+i in room_bookings[venue][day_idx] or
                                schedule_grid[day_idx][start_period+i]['type'] is not None or
                                is_rest_period(TIME_PERIODS[start_period+i])):
                                is_available = False
                                break
                        
                        if is_available:
                            # Mark professor and classroom as busy
                            for i in range(LAB_BLOCKS):
                                teacher_bookings[instructor][day_idx].add(start_period+i)
                                room_bookings[venue][day_idx].add(start_period+i)
                                schedule_grid[day_idx][start_period+i]['type'] = 'LAB'
                                schedule_grid[day_idx][start_period+i]['code'] = code_id if i == 0 else ''
                                schedule_grid[day_idx][start_period+i]['name'] = subj_name if i == 0 else ''
                                schedule_grid[day_idx][start_period+i]['faculty'] = instructor if i == 0 else ''
                                schedule_grid[day_idx][start_period+i]['classroom'] = venue if i == 0 else ''
                            is_scheduled = True
                    try_count += 1
            
            # Process lectures and tutorials
            theory_subjects = section_subjects[section_subjects['P'] == 0]
            for _, subject in theory_subjects.iterrows():
                code_id = str(subject['Course Code'])
                subj_name = str(subject['Course Name'])
                instructor = str(subject['Faculty'])
                venue = str(subject['Classroom'])
                lecture_hours = int(subject['L']) if pd.notna(subject['L']) else 0
                tutorial_hours = int(subject['T']) if pd.notna(subject['T']) else 0
                
                # Assign a color to this course if not already assigned
                if code_id not in subject_colors:
                    subject_colors[code_id] = {"color": next(color_generator), "name": subj_name, "faculty": instructor}
                
                if instructor not in teacher_bookings:
                    teacher_bookings[instructor] = {day_idx: set() for day_idx in range(len(WEEKDAYS))}
                if venue not in room_bookings:
                    room_bookings[venue] = {day_idx: set() for day_idx in range(len(WEEKDAYS))}
                
                # Schedule lectures (1.5 hours)
                for _ in range(lecture_hours):
                    is_scheduled = False
                    try_count = 0
                    while not is_scheduled and try_count < 1000:
                        day_idx = random.randint(0, len(WEEKDAYS)-1)
                        if len(TIME_PERIODS) >= LECTURE_BLOCKS:
                            start_period = random.randint(0, len(TIME_PERIODS)-LECTURE_BLOCKS)
                            
                            # Check if all required slots are free and not in break time
                            is_available = True
                            for i in range(LECTURE_BLOCKS):
                                if (start_period+i in teacher_bookings[instructor][day_idx] or 
                                    start_period+i in room_bookings[venue][day_idx] or
                                    schedule_grid[day_idx][start_period+i]['type'] is not None or
                                    is_rest_period(TIME_PERIODS[start_period+i])):
                                    is_available = False
                                    break
                            
                            if is_available:
                                # Mark professor and classroom as busy
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
                
                for period_idx in range(len(TIME_PERIODS)):
                    if skip_cells > 0:
                        skip_cells -= 1
                        continue
                    
                    # Count consecutive break periods
                    if is_rest_period(TIME_PERIODS[period_idx]):
                        break_count = 1
                        next_idx = period_idx + 1
                        while next_idx < len(TIME_PERIODS) and is_rest_period(TIME_PERIODS[next_idx]):
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
            
            # Add legend with improved formatting
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