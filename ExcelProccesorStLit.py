import streamlit as st
from ExcelProcessor import ExcelProcessor
from io import BytesIO

def main():
    st.title("Excel Processing App")
    st.sidebar.title("Input Parameters")

    # File uploader for Excel file
    uploaded_file = st.sidebar.file_uploader("Upload Excel file", type=["xlsx"])

    if uploaded_file:
        st.sidebar.header("Cutoff Percentages")
        cutoff_10th = st.sidebar.slider("10th Percentage Cutoff", min_value=0, max_value=100, value=70)
        cutoff_12th = st.sidebar.slider("12th Percentage Cutoff", min_value=0, max_value=100, value=70)
        cutoff_btech_cgpa = st.sidebar.slider("BTech CGPA Cutoff", min_value=0.0, max_value=10.0, value=6.0)
        cutoff_live_kt = st.sidebar.slider("Live KT Cutoff", min_value=0, max_value=10, value=0)
        cutoff_drop = st.sidebar.slider("Drop Cutoff", min_value=0, max_value=10, value=0)
        cutoff_gap = st.sidebar.slider("Gap Cutoff", min_value=0, max_value=10, value=2)

        if st.sidebar.button("Process"):
            # Instantiate ExcelProcessor with input parameters
            processor = ExcelProcessor(uploaded_file, cutoff_10th, cutoff_12th, cutoff_btech_cgpa,
                                       cutoff_live_kt, cutoff_drop, cutoff_gap)

            # Process the Excel file
            processor.clean_data()
            removed_data = processor.check_data()
            processor.remove_columns()
            processor.basic_editing()
            ask_percentage = True  # Set to False if CGPA is to be kept
            processor.remove_unwanted_data(ask_percentage)
            preferred_order = [
                'Full Name', 'Personal Email ID', 'CollegeRollNo', 'Contact No', 'Gender',
                'Branch', 'BTech Major Course', 'College Name', '10th Percentage',
                '10th Year of Passing', '12th/Diploma', '12th/Diploma Percentage',
                '12th/ Diploma Year of Passing', 'Degree', 'BTech CGPA', 'BTech Percentage',
                'Batch', 'Resume'
            ]
            processor.sort_columns(preferred_order)
            processor.add_serial_column()
            processor.format_sheet()
            processor.rename_sheet('RAIT')
            processor.fill_empty_cells_with_na()
            # processor.adjust_column_widths()

            # Save processed data and removed data
            processed_file_path = "processed_file.xlsx"
            processor.save_data(processed_file_path, removed_data)
            file = processor.adjust_column_widths(processed_file_path, "processed_file.xlsx").save("processed_file.xlsx")
            st.success("Processing completed successfully!")

            # Provide download link for processed file
            # st.markdown(f"Download processed file [here](file)", unsafe_allow_html=True)
            file_path = 'processed_file.xlsx'
            with open(file_path, 'rb') as file:
                file_data = file.read()
            file_buffer = BytesIO(file_data)

            # Provide a download button
            st.download_button(
                label="Download processed file",
                data=file_buffer,
                file_name='processed_with_widths.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )


if __name__ == "__main__":
    main()
