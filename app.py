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

            days = days_mapping.get(day.strip().upper(), [])

            for day in days:
                start_date = row['Start Date']
                end_date = row['End Date']
                current_date = start_date

                while current_date <= end_date:
                    if current_date.strftime('%A') == day:
                        if (current_date, start_time, end_time) not in added_dates:
                            start_datetime = datetime.combine(current_date, parse_time(start_time))
                            end_datetime = datetime.combine(current_date, parse_time(end_time))
                            add_event(row, start_datetime, end_datetime, clean_location(location))
                            added_dates.add((current_date, start_time, end_time))
                    current_date += timedelta(days=1)

    try:
        with open(filename, 'wb') as f:
            f.write(cal.to_ical())
        # st.write(f"ICS file created successfully: {filename}")
        st.write(f"Success!")
    except Exception as e:
        st.error(f"Failed to create ICS file: {e}")

def main():
    
    st.title('Carleton Calendar Converter')
    
    st.write("It's basically a tool that allows you to easily convert your Carleton academic schedule from an Excel `.xlsx` file to an `.ics` Calendar file.")
    
    st.image("img/main.gif")

    uploaded_file = st.file_uploader("Upload your Carleton course schedule (.xlsx)", type=["xlsx"])

    if uploaded_file:
        data = process_excel(uploaded_file)
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
        st.write("1. **Download Your Schedule as an Excel `.xlsx` file on Workday**: Go to Academics and Registration -> Registration Planning -> View My Courses & Saved Schedules")
        st.image("img/DownloadExcel/1.gif")
        st.write("2. **Upload Your Schedule**: Choose your Excel `.xlsx` file that contains your course schedule.")
        st.write("3. **Generate the Calendar**: Click `Download the .ics Calendar File`, which processes your file and generates an `.ics` file.")
        st.write("4. **Download and Import**: Download the `.ics` file and import it into your Apple Calendar.")
        st.image("img/HowItWorks/1.gif")
        st.write("5. **Voila**: Click on The Download and Press Ok.")
        st.image("img/HowItWorks/2.gif")

    elif selected == "Google Calendar":
        st.write("## How It Works / Tutorial for Google Calendar")
        st.write("1. **Download Your Schedule as an Excel `.xlsx` file on Workday**: Go to Academics and Registration -> Registration Planning -> View My Courses & Saved Schedules")
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
