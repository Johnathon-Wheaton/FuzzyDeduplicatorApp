import streamlit as st
import pandas as pd
import numpy as np
from thefuzz import fuzz
import io
import itertools
from collections import defaultdict

def calculate_similarity(row1, row2):
    """Calculate Jaro-Winkler similarity between two rows."""
    # Concatenate all fields into strings
    str1 = ' '.join(str(val) for val in row1 if pd.notna(val))
    str2 = ' '.join(str(val) for val in row2 if pd.notna(val))
    return fuzz.token_set_ratio(str1, str2) / 100.0

def get_comparison_groups(df, leading_chars):
    """Group rows by their leading characters to reduce comparison space."""
    # Concatenate all fields and get leading characters
    concatenated = df.apply(lambda row: ' '.join(str(val) for val in row if pd.notna(val)), axis=1)
    
    # Group by leading characters
    groups = defaultdict(list)
    for idx, text in enumerate(concatenated):
        if len(text) >= leading_chars:
            key = text[:leading_chars].lower()
            groups[key].append(idx)
    
    # Return only groups with more than one record
    return {k: v for k, v in groups.items() if len(v) > 1}

def calculate_total_comparisons(groups):
    """Calculate total number of comparisons needed given the grouping."""
    total = 0
    for indices in groups.values():
        n = len(indices)
        total += (n * (n - 1)) // 2
    return total

def find_duplicates(df, threshold=0.9, leading_chars=1):
    """Find duplicate rows using fuzzy matching."""
    # First group by leading characters
    groups = get_comparison_groups(df, leading_chars)
    
    duplicate_groups = []
    processed = set()
    
    # Create progress bar
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    total_comparisons = calculate_total_comparisons(groups)
    comparison_count = 0
    
    # Process each group separately
    for group_indices in groups.values():
        local_duplicates = set()
        
        for i, idx1 in enumerate(group_indices):
            if idx1 in processed:
                continue
                
            current_group = {idx1}
            row_i = df.iloc[idx1]
            
            for idx2 in group_indices[i + 1:]:
                if idx2 in processed:
                    continue
                    
                row_j = df.iloc[idx2]
                similarity = calculate_similarity(row_i, row_j)
                
                comparison_count += 1
                if comparison_count % 100 == 0:
                    progress = min(comparison_count / total_comparisons, 1.0)
                    progress_bar.progress(progress)
                    status_text.text(f"Processed {comparison_count:,} of {total_comparisons:,} comparisons...")
                
                if similarity >= threshold:
                    current_group.add(idx2)
                    processed.add(idx2)
            
            if len(current_group) > 1:
                duplicate_groups.append(current_group)
            processed.add(idx1)
    
    progress_bar.progress(1.0)
    status_text.text("Duplicate detection complete!")
    return duplicate_groups

def main():
    st.title("Excel Fuzzy Deduplication Tool")

    with st.expander("ℹ️ Instructions"):
        st.markdown("""
        ### What This Tool Does
        This tool helps you find and group similar records in your data, even when they're not exact matches. It:
        1. Considers **all columns** in your data for comparison
        2. Uses fuzzy matching to detect similar text (95% similarity threshold)
        3. Groups similar records together
        4. Provides an Excel file with all records labeled by group
        
        ### How It Works
        1. The tool combines all columns for each record into a single text string
        2. It compares every record with every other record using the Jaro-Winkler similarity algorithm
        3. Records with 95% or higher similarity are grouped together
        4. Results are provided in an Excel file with a 'Group_ID' column:
           - Records marked as 'Unique' have no similar matches
           - Records marked as 'Group_1', 'Group_2', etc. are similar to each other
        
        ### How to Use
        1. Copy your data from Excel or any spreadsheet
        2. Paste it into the text area below
        3. Wait for processing to complete
        4. Download the Excel file with the results
        
        ### Data Requirements
        - Include headers in the first row
        - Data can be tab-separated, comma-separated, or semicolon-separated
        - All columns will be used in the similarity comparison
        - Larger datasets will take longer to process
        """)
    
    # File upload
    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Read Excel file
        df = pd.read_excel(uploaded_file)
        st.write(f"Loaded {len(df):,} rows and {len(df.columns)} columns")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Similarity threshold slider
            threshold = st.slider(
                "Select similarity threshold",
                min_value=0.5,
                max_value=1.0,
                value=0.9,
                step=0.05,
                help="Higher values mean records need to be more similar to be considered duplicates"
            )
        
        with col2:
            # Leading characters slider
            leading_chars = st.slider(
                "Number of leading characters to match",
                min_value=1,
                max_value=10,
                value=3,
                step=1,
                help="Records must match on this many leading characters to be compared. Higher values = faster processing but might miss some duplicates"
            )
        
        # Calculate and show number of comparisons
        groups = get_comparison_groups(df, leading_chars)
        total_comparisons = calculate_total_comparisons(groups)
        total_possible = (len(df) * (len(df) - 1)) // 2
        
        st.write(f"Will perform {total_comparisons:,} comparisons")
        st.write(f"(Reduced from {total_possible:,} possible comparisons - {(total_comparisons/total_possible*100):.1f}% of original)")
        
        if st.button("Find Duplicates"):
            # Find duplicates
            duplicate_groups = find_duplicates(df, threshold, leading_chars)
            
            # Create new columns for duplicate information
            df['duplicate_group'] = -1
            df['duplicate_rows'] = ''
            
            # Populate duplicate information
            for group_idx, group in enumerate(duplicate_groups):
                group_list = sorted(group)
                rows_str = ', '.join(str(idx + 1) for idx in group_list)  # Adding 1 for Excel-style row numbers
                
                for idx in group_list:
                    df.at[idx, 'duplicate_group'] = group_idx + 1
                    df.at[idx, 'duplicate_rows'] = rows_str
            
            # Show summary
            n_duplicates = sum(len(group) for group in duplicate_groups)
            st.write(f"Found {n_duplicates:,} records in {len(duplicate_groups):,} duplicate groups")
            
            # Create download button
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            
            output.seek(0)
            st.download_button(
                label="Download Results",
                data=output,
                file_name="deduplication_results.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # Show preview of results
            st.write("Preview of results (showing first 1000 rows):")
            st.dataframe(df.head(1000))

if __name__ == "__main__":
    main()