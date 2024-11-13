import pypdf
from datetime import datetime, timedelta
from ics import Calendar, Event
import re
from zoneinfo import ZoneInfo

def read_pdf(pdf_path):
    """Read text directly from PDF file."""
    text = ""
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = pypdf.PdfReader(file)
            for page in pdf_reader.pages:
                text += page.extract_text() + "\n"
        return text
    except Exception as e:
        print(f"Error reading PDF: {e}")
        return None

def parse_time(time_str):
    """Parse time string in format 'H:MM' or 'HH:MM'."""
    try:
        if ':' not in time_str:
            return None
        hours, minutes = map(int, time_str.strip().split(':'))
        return f"{hours:02d}:{minutes:02d}"
    except:
        return None

def extract_first_shift_time(line):
    """Extract first shift information from a line."""
    try:
        # Regular expression to match first shift pattern
        # Match pattern: (BO|DP|DPO) followed by time, then service number and location
        pattern = r'(?:BO|DP|DPO)\s+(\d{1,2}:\d{2})\s+\d{1,2}:\d{2}\s+\w+\s+(\d+)\s+\d+\s+-\s+\d+\s+(\w+)'
        match = re.search(pattern, line)

        if match:

            print("MATCH FOR FIRST")
            start_time = parse_time(match.group(1))
            ldebpce = match.group(3)

            print(f"Found first shift time: Ligne {ldebpce} -> ({start_time})")

            return {
                'ldebpce': ldebpce,
                'start_time': start_time,
            }
    except Exception as e:
        print(f"Error extracting first shift info: {e}")
    return None

def extract_shifts_info(line):
    """Extract all subsequent shift information from a line."""
    shifts = []
    try:
        # Pattern for subsequent shifts
        pattern = r'(\d+)\s+(\d+)\s+-\s+(\d+)\s+(\w+)\s+(\d{1,2}:\d{2})\s+(\d{1,2}:\d{2})\s+(\w+)'
        matches = re.finditer(pattern, line)

        for match in matches:
            ligne = match.group(1)
            ldebpce = match.group(4)
            lfinpce = match.group(7)
            start_time = parse_time(match.group(5))
            end_time = parse_time(match.group(6))

            print(f"Found shift: Ligne {ligne} | {ldebpce} -> {lfinpce} ({start_time} - {end_time})")

            shifts.append({
                'ligne': ligne,
                'ldebpce': ldebpce,
                'lfinpce': lfinpce,
                'start_time': start_time,
                'end_time': end_time,
                'full_line': line
            })
    except Exception as e:
        print(f"Error extracting shift info: {e}")
    return shifts

def parse_schedule(text):
    """Parse the schedule text and return a list of shifts."""
    shifts = []
    current_date = None

    # Split into lines and remove empty lines
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    first_shift_time = None

    for line in lines:
        # Skip header lines and footer
        if any(x in line for x in ['tl Jour', 'HASTUS', 'Division:', 'Page:']):
            continue

        # Check for new date line
        if line.startswith('9387'):
            date_match = re.search(r'9387\s+\w+\.\s+(\d{2}/\d{2}/\d{4})', line)
            if date_match:
                current_date = date_match.group(1)
                print(f"\nProcessing date: {current_date}")

                # Skip rest days and holidays
                if any(x in line.split()[3:5] for x in ['R', 'C', 'VAP', 'X']):
                    print("Skipping rest/holiday day")
                    continue

                # Extract first shift from this line
                first_shift_time = extract_first_shift_time(line)


                # Extract any additional shifts from the same line
                second_shifts = extract_shifts_info(line)
                for shift_info in second_shifts:
                    if second_shifts.index(shift_info) == 0 and first_shift_time:
                        shift_info['ldebpce'] = first_shift_time['ldebpce']
                        shift_info['start_time'] = first_shift_time['start_time']
                        first_shift_time = None

                    shift = create_shift(current_date, shift_info)
                    if shift:
                        shifts.append(shift)

        # Process continuation lines
        elif current_date and '-' in line:
            second_shifts = extract_shifts_info(line)
            for shift_info in second_shifts:
                if second_shifts.index(shift_info) == 0 and first_shift_time:
                    shift_info['ldebpce'] = first_shift_time['ldebpce']
                    shift_info['start_time'] = first_shift_time['start_time']
                    first_shift_time = None
                shift = create_shift(current_date, shift_info)
                if shift:
                    shifts.append(shift)

    return shifts

def create_shift(date_str, shift_info):
    """Create a shift dictionary from parsed information."""
    try:
        zurich_tz = ZoneInfo("Europe/Zurich")

        start_dt = datetime.strptime(f"{date_str} {shift_info['start_time']}", '%d/%m/%Y %H:%M')
        end_dt = datetime.strptime(f"{date_str} {shift_info['end_time']}", '%d/%m/%Y %H:%M')

        # Add timezone
        start_dt = start_dt.replace(tzinfo=zurich_tz)
        end_dt = end_dt.replace(tzinfo=zurich_tz)

        # If end time is before start time, add one day to end time
        if end_dt < start_dt:
            end_dt += timedelta(days=1)

        duration = end_dt - start_dt
        duration_str = f"{duration.total_seconds() / 3600:.2f}h"

        return {
            'date': start_dt.date(),
            'start': start_dt,
            'end': end_dt,
            'ligne': shift_info['ligne'],
            'ldebpce': shift_info['ldebpce'],
            'lfinpce': shift_info['lfinpce'],
            'duration': duration_str,
            'details': shift_info['full_line']
        }
    except Exception as e:
        print(f"Error creating shift: {e}")
        return None

def create_calendar(shifts):
    """Create an ICS calendar from the list of shifts."""
    cal = Calendar()

    # Sort shifts by start time
    shifts.sort(key=lambda x: x['start'])

    # Add events for shifts
    for shift in shifts:
        event = Event()
        event.name = f"Ligne {shift['ligne']} | {shift['ldebpce']} -> {shift['lfinpce']}"
        event.begin = shift['start']
        event.end = shift['end']
        event.location = f"{shift['ldebpce']} -> {shift['lfinpce']}"
        event.description = (
            f"Work shift details:\n"
            f"Duration: {shift['duration']}\n"
            f"Ligne: {shift['ligne']}\n"
            f"From: {shift['ldebpce']}\n"
            f"To: {shift['lfinpce']}\n"
            f"Full details: {shift['details']}"
        )
        cal.events.add(event)

    return cal

def convert_pdf_to_ics(pdf_path, output_file):
    """Convert PDF schedule to ICS file."""
    # Read the PDF
    text = read_pdf(pdf_path)
    if not text:
        print("Error: Could not read PDF file")
        return 0

    # Parse the schedule
    shifts = parse_schedule(text)

    if not shifts:
        print("Error: No shifts found in the schedule")
        return 0

    # Create calendar
    cal = create_calendar(shifts)

    # Write to file
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(str(cal))

    return len(shifts)

if __name__ == "__main__":
    pdf_file = 'schedule.pdf'
    ics_file = 'work_schedule.ics'

    num_shifts = convert_pdf_to_ics(pdf_file, ics_file)

    if num_shifts > 0:
        print(f"\nSuccessfully created calendar with {num_shifts} shifts.")
    else:
        print("Failed to create calendar. Please check the PDF file and error messages.")
