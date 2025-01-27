import pandas as pd
from datetime import datetime
from collections import defaultdict

# File paths for the uploaded Excel files
trainers_file = 'personal_trainers.xlsx'  # Updated to match your uploaded file
clients_file = 'clients.xlsx'  # Updated to match your uploaded file
output_file = 'PT_Client_Matching_Output.xlsx'  # Save output locally

# Read the Excel files
trainers_df = pd.read_excel(trainers_file)
clients_df = pd.read_excel(clients_file)

def parse_day_time(entry):
    if pd.isna(entry):
        return []  # Return an empty list if the entry is NaN

    ranges = []
    segments = [segment.strip() for segment in entry.split(",")]

    for segment in segments:
        # Standardize ranges
        segment = segment.replace("all day", "6am-10pm").replace(" to ", "-")

        if " " not in segment:
            continue

        try:
            days_part, time_ranges = segment.split(" ", 1)
            days = [day.strip() for day in days_part.split("/")]
            time_ranges = [time.strip() for time in time_ranges.replace("&", ",").split(",")]

            for time_range in time_ranges:
                # Standardize time range
                if "-" in time_range and ("am" in time_range or "pm" in time_range):
                    parts = time_range.split("-")
                    suffix = "am" if "am" in parts[1] else "pm"
                    if not parts[0].endswith(("am", "pm")):
                        parts[0] += suffix
                    time_range = f"{parts[0]}-{parts[1]}"

                if "-" not in time_range:
                    continue

                start_time, end_time = [normalize_time_format(t.strip()) for t in time_range.split("-")]

                if start_time and end_time:
                    ranges.extend({'day': day, 'start_time': start_time, 'end_time': end_time} for day in days)
        except ValueError:
            continue

    return ranges

# Helper function to normalize time format
def normalize_time_format(time_str):
    try:
        if ":" in time_str:
            return datetime.strptime(time_str, "%I:%M%p").strftime("%I:%M%p")
        else:
            return datetime.strptime(time_str, "%I%p").strftime("%I:%M%p")
    except ValueError:
        return None

# Helper function to check time overlap and calculate overlap duration
def get_overlap_duration(client_start, client_end, trainer_start, trainer_end):
    fmt = "%I:%M%p"
    try:
        client_start = datetime.strptime(client_start, fmt)
        client_end = datetime.strptime(client_end, fmt)
        trainer_start = datetime.strptime(trainer_start, fmt)
        trainer_end = datetime.strptime(trainer_end, fmt)
    except ValueError:
        return 0

    overlap_start = max(client_start, trainer_start)
    overlap_end = min(client_end, trainer_end)
    if overlap_start < overlap_end:
        return (overlap_end - overlap_start).seconds / 3600
    return 0

def create_availability_dict(dataframe, is_client=False):
    availability_dict = defaultdict(list)

    for _, row in dataframe.iterrows():
        # Clean and standardize key fields
        name = str(row.get('Name', 'Unknown')).strip()
        availability = str(row.get('Availability', '')).strip()
        location = str(row.get('Location', 'Unknown')).strip()
        notes = str(row.get('Notes', '')).strip().lower() if not is_client else ''  # Skip notes for clients

        # Skip rows with missing or invalid 'Name'
        if not name or name.lower() == 'unknown':
            continue

        gender_preference = "N/A"  # Default value for gender preference
        if is_client:
            gender_preference = str(row.get('Gender Preference', 'N/A')).strip().lower()

        gender = "N/A"  # Default value for trainer gender
        if not is_client:
            gender = str(row.get('Gender', 'N/A')).strip().lower()

        # Parse and standardize availability ranges
        availability_ranges = parse_day_time(availability)
        for range_entry in availability_ranges:
            entry = {
                'name': name,
                'start_time': range_entry['start_time'],
                'end_time': range_entry['end_time'],
                'Location': location,
                'Notes': notes if not is_client else None,
            }
            if is_client:
                entry['Gender Preference'] = gender_preference
            else:
                entry['Gender'] = gender
            availability_dict[range_entry['day']].append(entry)

    return availability_dict

# Function to process matches for a single day
def process_day_matches(day, client_entries, trainer_entries, client_matches):
    for client_entry in client_entries:
        gender_preference = client_entry.get('Gender Preference', 'N/A').strip().lower()

        for trainer_entry in trainer_entries:
            trainer_gender = trainer_entry.get('Gender', 'nan').strip().lower()

            # Include all trainers for clients with "N/A" preference, otherwise match gender
            if gender_preference == "nan" or gender_preference == trainer_gender:
                # Check location match
                if (client_entry['Location'].strip() == trainer_entry['Location'].strip() or
                        trainer_entry['Location'].strip() == "Either"):
                    # Calculate overlap
                    overlap_duration = get_overlap_duration(
                        client_entry['start_time'], client_entry['end_time'],
                        trainer_entry['start_time'], trainer_entry['end_time']
                    )
                    if overlap_duration > 0:
                        client_matches[client_entry['name']][trainer_entry['name']] += overlap_duration

# Create trainer and client dictionaries
trainer_dict = create_availability_dict(trainers_df, is_client=False)
client_dict = create_availability_dict(clients_df, is_client=True)

# Initialize a dictionary to store matches by client
client_matches = defaultdict(lambda: defaultdict(float))

# Main matching logic
for day, client_entries in client_dict.items():
    if day in trainer_dict:
        trainer_entries = trainer_dict[day]
        process_day_matches(day, client_entries, trainer_entries, client_matches)

# Print the output for each client with trainers and total overlap durations
for client_name, trainers in client_matches.items():
    print(f"\nClient: {client_name}")
    for trainer_name, total_overlap in sorted(trainers.items(), key=lambda x: x[1], reverse=True):
        print(f"  Trainer: {trainer_name}, Total Overlap: {total_overlap:.1f} hours")

# Save only the top 5 trainers for each client to the Excel file
output_rows = []
for client, trainers in client_matches.items():
    top_5_trainers = sorted(trainers.items(), key=lambda x: x[1], reverse=True)[:5]
    formatted_trainers = ', '.join([f"{trainer} ({overlap:.1f}h)" for trainer, overlap in top_5_trainers])
    output_rows.append({'Client Name': client, 'Top 5 Trainers': formatted_trainers})

output_df = pd.DataFrame(output_rows)
output_df.to_excel(output_file, index=False)
print(f"Matching complete. Output saved to {output_file}")