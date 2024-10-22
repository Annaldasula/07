import streamlit as st
import pandas as pd
import os
from io import BytesIO  # Needed to handle file as bytes

# Function to extract entity name from file path
def extract_entity_name(file_path):
    base_name = os.path.basename(file_path)
    entity_name = base_name.split('_or_')[0].replace("_", " ").split('-')[0].strip()
    return entity_name

# Web app title
st.title('Excel File Merger & Entity Extractor')

# Initialize session state for the file uploader
if 'uploaded_files' not in st.session_state:
    st.session_state.uploaded_files = []

# File uploader
uploaded_files = st.file_uploader("Upload your Excel files", accept_multiple_files=True, type=['xlsx'])

if uploaded_files:
    # Store uploaded files in session state
    st.session_state.uploaded_files = uploaded_files

if st.session_state.uploaded_files:
    final_df = pd.DataFrame()
    
    # Loop through each uploaded file
    for uploaded_file in st.session_state.uploaded_files:
        df = pd.read_excel(uploaded_file)
        
        # Extract the entity name and add it as a new column
        entity_name = extract_entity_name(uploaded_file.name)
        df['Entity'] = entity_name
        
        # Concatenate all the dataframes
        final_df = pd.concat([final_df, df], ignore_index=True)
    
    # Process columns as required
    existing_columns = final_df.columns.tolist()
    influencer_index = existing_columns.index('Influencer')
    country_index = existing_columns.index('Country')
    
    new_order = (
        existing_columns[:influencer_index + 1] +  # All columns up to and including 'Influencer'
        ['Entity', 'Reach', 'Sentiment', 'Keywords', 'State', 'City', 'Engagement'] +  # Adding new columns
        existing_columns[influencer_index + 1:country_index + 1]  # All columns between 'Influencer' and 'Country'
    )
    
    # Fill missing values in 'Influencer' column with 'Bureau News'
    final_df['Influencer'] = final_df['Influencer'].fillna('Bureau News')
    final_df['Date'] = pd.to_datetime(final_df['Date']).dt.normalize()
    
    # Reorder the DataFrame
    final_df = final_df[new_order]
    
    # Show the processed dataframe in the web app
    st.write(final_df)
    
    # Prepare Excel file in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        final_df.to_excel(writer, index=False)
    
    # Convert buffer to bytes
    processed_data = output.getvalue()

    # Option to download the merged file
    st.download_button(
        label="Download Merged Excel",
        data=processed_data,
        file_name='merged_excel_with_entity.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

# Button to clear uploaded files
if st.button('Clear Uploads'):
    st.session_state.uploaded_files = []
    st.experimental_rerun()  # Reload the app to reflect the change
