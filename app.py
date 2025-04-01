
import streamlit as st
import pandas as pd
import re

st.set_page_config(page_title="Coast2coast X Red Cross Online Course Report Generator")

st.title("Coast2coast X Red Cross Online Course Report Generator")

# Helper functions
def extract_location(course_type):
    match = re.search(r'\(([^)]+)\)', str(course_type))
    return match.group(1).strip() if match else None

def filter_valid_courses(df):
    valid_levels = {
        "Standard First Aid & CPR/AED Level C": "SFA",
        "Emergency First Aid CPR/AED Level C": "EFA",
        "CPR/AED Level C": "CPR",
        "Marine Basic First Aid & CPR/AED Level C": "MBFA"
    }
    df = df[df['Courses & Levels'].isin(valid_levels.keys())].copy()
    df['Course Level Code'] = df['Courses & Levels'].map(valid_levels)
    df['Location'] = df['COURSE TYPE'].apply(extract_location)
    return df

def generate_red_cross_upload(df, course_id_df):
    output_rows = []
    unmatched = []
    for _, row in df.iterrows():
        match = course_id_df[
            (course_id_df['Start Date'] == row['Start']) &
            (course_id_df['Facility'].str.endswith(f"- {row['Location']}", na=False)) &
            (course_id_df['Course Level'].str.contains(row['Course Level Code'], case=False, na=False))
        ]
        if len(match) == 1:
            course_id = match.iloc[0]['Course ID']
            output_rows.append({
                "Course Number No du cours": course_id,
                "First Name Prénom": row['First name (participant)'],
                "Last Name Nom de famille": row['Last name (participant)'],
                "Email Courriel": row['Email address (participant)']
            })
        else:
            unmatched.append({
                "First Name": row['First name (participant)'],
                "Last Name": row['Last name (participant)'],
                "Email": row['Email address (participant)'],
                "Date": row['Start'],
                "Location": row['Location'],
                "Course Level": row['Courses & Levels'],
                "Reason": "No matching course ID" if len(match) == 0 else "Duplicate matches"
            })
    return pd.DataFrame(output_rows), pd.DataFrame(unmatched)

def generate_upsell_list(df):
    upsell_levels = {"EFA", "CPR"}
    upsell_df = df[df['Course Level Code'].isin(upsell_levels)].copy()
    return upsell_df[[
        'First name (participant)', 'Last name (participant)',
        'Email address (participant)', 'Phone (participant)',
        'Location', 'Courses & Levels'
    ]]

# File uploads
bookeo_file = st.file_uploader("Upload Bookeo Report", type=["xlsx"])
course_id_file = st.file_uploader("Upload Course ID List", type=["xlsx"])

if bookeo_file and course_id_file:
    bookeo_df = pd.read_excel(bookeo_file)
    course_df = pd.read_excel(course_id_file)

    try:
        bookeo_df_filtered = filter_valid_courses(bookeo_df)

        red_cross_df, unmatched_df = generate_red_cross_upload(bookeo_df_filtered, course_df)
        upsell_df = generate_upsell_list(bookeo_df_filtered)

        st.success("Files processed successfully!")

        st.download_button("Download Red Cross Upload File", red_cross_df.to_csv(index=False), "red_cross_upload.csv")
        st.download_button("Download Upsell Call List", upsell_df.to_csv(index=False), "upsell_call_list.csv")
        st.download_button("Download Unmatched Students Report", unmatched_df.to_csv(index=False), "unmatched_students.csv")

    except Exception as e:
        st.error(f"An error occurred: {e}")
