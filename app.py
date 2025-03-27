import streamlit as st
from streamlit_option_menu import option_menu
import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime, timedelta

# Title for browser tab
st.set_page_config(
    page_title="Carleton Calendar Converter",  
    page_icon="📅",  
    layout="centered",  
    initial_sidebar_state="auto" 
)

def process_excel(file):
    try:
        # Read the Excel file with header=None to get raw data
        df = pd.read_excel(file, header=None)
        
        # Find the index of "My Enrolled Courses"
        enrolled_index = None
        for idx, row in df.iterrows():
            if row[0] == "My Enrolled Courses":
                enrolled_index = idx
                break
                
        if enrolled_index is None:
            st.error("Could not find 'My Enrolled Courses' in the file.")
            return None
            
        # Find the index of "My Waitlisted Courses" or any other endpoint
        end_index = None
        for idx, row in df.iloc[enrolled_index+1:].iterrows():
            if pd.notna(row[0]) and row[0] in ["My Waitlisted Courses", "My Completed Courses", "My Dropped/Withdrawn Courses"]:
                end_index = idx
                break
                
        if end_index is None:
            st.error("Could not find the end of enrolled courses section.")
            return None
            
        # Find the header row which is typically 2 rows after "My Enrolled Courses"
        header_index = enrolled_index + 2
        
        # Extract relevant data between the header row and the end index
        data_start = header_index + 1
        relevant_df = df.iloc[data_start:end_index].copy()
        
        # Get column indices for the required fields
        # Examining the Excel data, the columns are:
        # Section at index 5
        # Meeting Patterns at index 9
        # Start Date at index 11
        # End Date at index 12
        relevant_df = relevant_df.iloc[:, [5, 9, 11, 12]]
        
        # Set column headers
        relevant_df.columns = ['Section', 'Meeting Patterns', 'Start Date', 'End Date']
        
        # Convert date columns to datetime
        relevant_df['Start Date'] = pd.to_datetime(relevant_df['Start Date'], errors='coerce')
        relevant_df['End Date'] = pd.to_datetime(relevant_df['End Date'], errors='coerce')
        
        # Drop rows with NaN dates
        relevant_df = relevant_df.dropna(subset=['Start Date', 'End Date'])
        
        return relevant_df
        
    except Exception as e:
        st.error(f"Error processing Excel file: {str(e)}")
        return None

def create_ics_file(events, filename='schedule.ics'):
    cal = Calendar()
    from icalendar import vRecur

    def parse_time(t):
        return datetime.strptime(t.strip(), '%I:%M %p').time()

    def add_recurring_event(event, first_start_datetime, first_end_datetime, location, recurrence_days, end_date):
        ics_event = Event()
        ics_event.add('summary', event['Section'])
        ics_event.add('location', location)
        ics_event.add('dtstart', first_start_datetime)
        ics_event.add('dtend', first_end_datetime)
        
        # Set recurrence rule
        recur_rule = vRecur()
        recur_rule['FREQ'] = 'WEEKLY'
        recur_rule['UNTIL'] = end_date.date()
        recur_rule['BYDAY'] = recurrence_days
        ics_event.add('rrule', recur_rule)
        
        cal.add_component(ics_event)

    def clean_location(location_str):
        return location_str.strip().replace('\n', ', ')
        
    # Define day mappings for recurrence rules
    day_to_rrule = {
        'Monday': 'MO',
        'Tuesday': 'TU',
        'Wednesday': 'WE', 
        'Thursday': 'TH',
        'Friday': 'FR'
    }

    for _, row in events.iterrows():
        processed_patterns = set()

        # Check if 'Meeting Patterns' is valid (not NaN and is a string)
        if pd.notna(row['Meeting Patterns']) and isinstance(row['Meeting Patterns'], str):
            for pattern in row['Meeting Patterns'].split('\n'):
                parts = pattern.split(' | ')
                if len(parts) < 3:
                    continue

                day_code, time_range, location = parts
                start_time, end_time = time_range.split(' - ')
                
                # Skip if we've already processed this pattern
                pattern_key = (day_code, time_range)
                if pattern_key in processed_patterns:
                    continue
                processed_patterns.add(pattern_key)

                days_mapping = {
                    'M': ['Monday'],
                    'T': ['Tuesday'],
                    'W': ['Wednesday'],
                    'TH': ['Thursday'],
                    'F': ['Friday'],
                    'MW': ['Monday', 'Wednesday'],
                    'TTH': ['Tuesday', 'Thursday'],
                    'MWF': ['Monday', 'Wednesday', 'Friday'],
                }

                days = days_mapping.get(day_code.strip().upper(), [])
                if not days:
                    continue
                
                # Convert days to RRULE format (MO,WE,FR)
                recurrence_days = [day_to_rrule[day] for day in days]
                
                # Find the first occurrence of each day to start the recurring event
                start_date = row['Start Date']
                end_date = row['End Date']
                current_date = start_date

                # Find the first date that matches one of our days
                first_date = None
                while current_date <= end_date and first_date is None:
                    if current_date.strftime('%A') in days:
                        first_date = current_date
                    current_date += timedelta(days=1)
                
                if first_date:
                    start_datetime = datetime.combine(first_date, parse_time(start_time))
                    end_datetime = datetime.combine(first_date, parse_time(end_time))
                    add_recurring_event(
                        row, 
                        start_datetime, 
                        end_datetime, 
                        clean_location(location),
                        recurrence_days,
                        end_date
                    )

    try:
        with open(filename, 'wb') as f:
            f.write(cal.to_ical())
        st.success("Calendar file created successfully!")
    except Exception as e:
        st.error(f"Failed to create ICS file: {e}")

def main():
    
    st.title('Carleton Calendar Converter')
    
    st.write("It's basically a tool that allows you to easily convert your Carleton academic schedule from an Excel `.xlsx` file to an `.ics` Calendar file.")
    
    # st.image("img/main.gif")

    uploaded_file = st.file_uploader("Upload your Carleton course schedule (.xlsx)", type=["xlsx"])

    if uploaded_file:
        with st.spinner("Processing your schedule..."):
            data = process_excel(uploaded_file)
            
        if data is not None and not data.empty:
            st.write("Your Courses for The Term:")
            st.write(data)
            
            create_ics_file(data)
            
            with open('schedule.ics', 'rb') as f:
                st.download_button(
                    label="Download the `.ics` Calendar File",
                    data=f,
                    file_name='myCarletonSchedule.ics',
                    mime='text/calendar'
                )
        else:
            st.error("Could not extract course data from the file. Please make sure you're uploading the correct Excel file.")

    st.markdown("---")
    selected = option_menu(
        menu_title="Tutorials", 
        options=["MacBook", "Google Calendar"],  
        icons=["apple", "google"],  
        menu_icon="book", 
        default_index=0,  
        orientation="horizontal",
    )

    if selected == "MacBook":
        st.write("## How It Works / Tutorial for MacBook")
        st.write("1. **Download Your Schedule as an Excel `.xlsx` file on Workday**: Go to `Academics and Registration` -> `Registration Planning` -> `View My Courses` (DONT CLICK View My Saved Schedules!)")
        st.image("img/DownloadExcel/1.gif")
        st.write("2. **Upload Your Schedule**: Choose your Excel `.xlsx` file that contains your course schedule.")
        st.write("3. **Generate the Calendar**: Click `Download the .ics Calendar File`, which processes your file and generates an `.ics` file.")
        st.write("4. **Download and Import**: Download the `.ics` file and import it into your Apple Calendar.")
        st.image("img/HowItWorks/1.gif")
        st.write("5. **Voila**: Click on The Download and Press Ok.")
        st.image("img/HowItWorks/2.gif")

    elif selected == "Google Calendar":
        st.write("## How It Works / Tutorial for Google Calendar")
        st.write("1. **Download Your Schedule as an Excel `.xlsx` file on Workday**: Go to `Academics and Registration` -> `Registration Planning` -> `View My Courses` (DONT CLICK View My Saved Schedules!)")
        st.image("img/DownloadExcel/1.gif")
        st.write("2. **Upload Your Schedule**: Choose your Excel `.xlsx` file that contains your course schedule.")
        st.write("3. **Generate the Calendar**: Click `Download the .ics Calendar File`, which processes your file and generates an `.ics` file.")
        st.write("4. **Download the `.ics` file**.")
        st.image("img/HowItWorks/1.gif")
        st.write('5. **Open Google Calendar**: In the top right, click on the Settings icon, then select "Settings."')
        st.image("img/HowItWorks/Google/1.jpg")
        st.write('6. **Click on Import & Export**: Within the Settings menu, locate and click on "Import & Export."')
        st.image("img/HowItWorks/Google/2.jpg")
        st.write('7. **Import the `.ics` file**: Click "Select file from your computer," and choose the `.ics` file you downloaded earlier.')
        st.image("img/HowItWorks/Google/3.jpg")
        st.write('8. **Choose Your Calendar**: Select which calendar to add the imported events to—by default, events will be imported into your primary calendar. Click "Import" to finalize.')
        st.image("img/HowItWorks/Google/4.jpg")
        st.write("9. **Voila**: Your course schedule is now successfully imported into your Google Calendar, ready for you to view and manage.")

    st.markdown("---")
    st.write('Disclaimer: This tool is provided "as-is" without any guarantees. The developer is not liable for any damages resulting from the use of this application. This project is not affiliated with Carleton College.')
    st.write("Feel free to reach out to me at gautamaj@carleton.edu for any inquiries.")
    st.write("Contribute to this project on [Github](https://github.com/JohnCassavetes/CarletonCalendarConverter).")

if __name__ == "__main__":
    main()