This Python script is designed to match clients with personal trainers based on their availability, location, and preferences. 
It reads two Excel files containing trainer and client data, processes their availability, 
and identifies the top 5 trainers for each client based on overlapping schedules and compatibility.

Features
Input Data Parsing: Reads client and trainer information from Excel files.
Availability Parsing: Processes complex availability formats (e.g., "Mon 9am-11am, Tue 1pm-3pm").
Matching Logic: Matches trainers and clients based on:
Overlapping time slots.
Location compatibility.
Gender preferences (if specified).
Output Generation: Saves the top 5 trainers for each client along with overlap details to a new Excel file.
Input Files
Trainers File: personal_trainers.xlsx

Columns: Name, Availability, Location, Gender, Notes
Clients File: clients.xlsx

Columns: Name, Availability, Location, Gender Preference
Notes:
Availability format: Days and times (e.g., "Mon 9am-5pm, Tue 10am-3pm").
Location: Indicates where sessions are preferred (e.g., Gym A, Gym B, Either).
Output File
File Name: PT_Client_Matching_Output.xlsx
Content:
Client Name: Name of the client.
Top 5 Trainers: Trainers with the highest overlap, formatted as Trainer Name (Overlap Duration).
How to Run
Install Required Libraries: Make sure Python is installed, along with the required libraries:

- pip install pandas openpyxl
Prepare Input Files: Ensure the personal_trainers.xlsx and clients.xlsx files are in the same directory as the script.

Run the Script: Execute the script in your terminal:

python PTMatching.py
View Results: Check the PT_Client_Matching_Output.xlsx file for the matching results.

Key Functions
parse_day_time(entry): Parses and standardizes day-time availability entries.

normalize_time_format(time_str): Ensures all time formats are consistent.

get_overlap_duration(client_start, client_end, trainer_start, trainer_end): Calculates the overlap duration between client and trainer availability.

create_availability_dict(dataframe, is_client): Creates dictionaries for trainer and client availability.

process_day_matches(day, client_entries, trainer_entries, client_matches): Processes matches for a single day based on availability and preferences.

Customization
Add Columns: Modify the script to include additional filters such as certifications or special skills.
Expand Output: Save detailed matching data (e.g., overlap times for all trainers) to the output file.
Example Output
For a client "John Doe," the output might look like:

Client Name	Top 5 Trainers
John Doe	Alice (5.0h), Bob (4.5h), Carol (4.0h)
