import streamlit as st
import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime, timedelta

def process_excel(file):
    df = pd.read_excel(file, header=None)
    
    # Identify the start index for "My Enrolled Courses"
    start_index = df[df.iloc[:, 0].str.contains('My Enrolled Courses', na=False)].index.min()
    if pd.isna(start_index):
        raise ValueError("Could not find 'My Enrolled Courses' in the file.")
    
    # Initialize list to collect relevant rows
    relevant_data = []
    
    # Process rows from start_index + 2 until we hit "My Dropped/Withdrawn Courses"
    for index, row in df.iloc[start_index + 2:].iterrows():
        if 'My Dropped/Withdrawn Courses' in str(row[0]):
            break
        relevant_data.append(row)
    
    # Create DataFrame from the collected data
    relevant_df = pd.DataFrame(relevant_data)
    
    # Select relevant columns and set column headers
    relevant_df = relevant_df.iloc[:, [4, 7, 10, 11]]
    relevant_df.columns = ['Section', 'Meeting Patterns', 'Start Date', 'End Date']
    
    # Convert date columns to datetime and drop rows with NaN dates
    relevant_df['Start Date'] = pd.to_datetime(relevant_df['Start Date'], format='%m/%d/%y', errors='coerce')
    relevant_df['End Date'] = pd.to_datetime(relevant_df['End Date'], format='%m/%d/%y', errors='coerce')
    relevant_df = relevant_df.dropna(subset=['Start Date', 'End Date'])

    return relevant_df

def create_ics_file(events, filename='schedule.ics'):
    cal = Calendar()
    
    def parse_time(t):
        return datetime.strptime(t.strip(), '%I:%M %p').time()
    
    def add_event(event, start_datetime, end_datetime, location):
        ics_event = Event()
        ics_event.add('summary', event['Section'])
        ics_event.add('location', location)
        ics_event.add('dtstart', start_datetime)
        ics_event.add('dtend', end_datetime)
        cal.add_component(ics_event)
    
    def clean_location(location_str):
        return location_str.strip().replace('\n', ', ')
    
    for _, row in events.iterrows():
        added_dates = set()  # Track added dates to avoid duplicates
        
        for pattern in row['Meeting Patterns'].split('\n'):
            parts = pattern.split(' | ')
            if len(parts) < 3:
                continue
            
            day, time_range, location = parts
            start_time, end_time = time_range.split(' - ')
            
            if 'MW' in day:
                days = ['Monday', 'Wednesday']
            elif 'TTH' in day:
                days = ['Tuesday', 'Thursday']
            elif 'F' in day:
                days = ['Friday']
            else:
                days = []
            
            for day in days:
                # Calculate the specific date for the day of the week
                start_date = row['Start Date']
                end_date = row['End Date']
                
                current_date = start_date
                while current_date <= end_date:
                    if current_date.strftime('%A') == day:
                        if (current_date, start_time, end_time) not in added_dates:
                            start_datetime = datetime.combine(current_date, parse_time(start_time))
                            end_datetime = datetime.combine(current_date, parse_time(end_time))
                            add_event(row, start_datetime, end_datetime, clean_location(location))
                            added_dates.add((current_date, start_time, end_time))  # Mark this event as added
                    current_date += timedelta(days=1)
    
    with open(filename, 'wb') as f:
        f.write(cal.to_ical())

def main():
    st.title('Carleton Calendar Converter')
    st.write('This tool converts your Carleton course schedule into an iCal file. (Currently only available for Apple Calendar)')
    st.write()
    uploaded_file = st.file_uploader("Upload your Carleton course schedule (.xlsx)", type=["xlsx"])

    if uploaded_file:
        data = process_excel(uploaded_file)
        st.write("Your Courses for The Term:")
        st.write(data)
        
        create_ics_file(data)
        
        with open('schedule.ics', 'rb') as f:
            st.download_button(
                label="Download The Apple Calendar File",
                data=f,
                file_name='myCarletonSchedule.ics',
                mime='text/calendar'
            )

    st.markdown("---")

    st.write("## Tutorial")

    st.write("1. Choose your Excel `.xlsx` file that contains your course schedule.")
    st.write("2. Click `Download the Apple Calendar File`. The app processes your file and generates an `.ics` file.")
    st.write("3. Download the `.ics` file and import it into your Apple Calendar.")
    st.image("img/HowItWorks/1.gif")
    st.write("4. Click on The Download and Press Ok.")
    st.image("img/HowItWorks/2.gif")

    st.write("### How To Download Your Schedule as an Excel `.xlsx` file")
    st.write("Go to Workday -> Academics and Registration -> Registration Planning -> View My Courses & Saved Schedules")
    st.image("img/DownloadExcel/1.gif")

    st.markdown("---")

    st.write("It basically works, but legally I have to add a... Disclaimer: The developer is not liable for any damages resulting from the use of this application.")
    st.write("Feel free to reach out to me at gautamaj@carleton.edu for any inquiries.")
    st.write("Contribute to this project on [Github](https://github.com/JohnCassavetes/CarletonCalendarConverter).")

if __name__ == "__main__":
    main()
