import streamlit as st
import pandas as pd
import base64
from io import BytesIO

# Import your existing function
from pms_visualization import generate_pms_visualization

def main():
    st.set_page_config(
        page_title="PMS Resource Analytics",
        page_icon="📊",
        layout="wide"
    )
    
    st.title("🚀 Enterprise Resource Analytics Platform")
    st.markdown("### Transform Excel data into interactive visualizations")
    
    # Sidebar for file upload
    st.sidebar.header("Upload Excel File")
    uploaded_file = st.sidebar.file_uploader(
        "Choose an Excel file", 
        type=['xlsx', 'xls'],
        help="Upload your PMS data file"
    )
    
    # Load sample data button
    if st.sidebar.button("Use Sample Data"):
        # You can include your Test_data.xlsx here
        st.info("Sample data loaded! (You'll need to include Test_data.xlsx in your repo)")
    
    if uploaded_file is not None:
        try:
            # Read the Excel file
            df = pd.read_excel(uploaded_file, engine='openpyxl')
            
            # Display data preview
            st.subheader("📋 Data Preview")
            st.dataframe(df.head())
            
            # Show data statistics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Records", len(df))
            with col2:
                st.metric("Unique Roles", df['Current Role'].nunique())
            with col3:
                st.metric("Regions", df['Region'].nunique())
            
            # Generate visualization
            if st.button("🎯 Generate Visualization", type="primary"):
                with st.spinner("Generating interactive visualization..."):
                    html_content = generate_pms_visualization(dataframe=df)
                    
                    # Display the HTML
                    st.subheader("📊 Interactive PMS Visualization")
                    st.components.v1.html(html_content, height=600, scrolling=True)
                    
                    # Download button
                    st.download_button(
                        label="💾 Download HTML",
                        data=html_content,
                        file_name="pms_visualization.html",
                        mime="text/html"
                    )
                    
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.info("Please ensure your Excel file has the required columns: Current Role, Region, Avail Bucket, Associate ID, Associate Name, Current Availability")
    
    else:
        st.info("👆 Please upload an Excel file to get started")
        
        # Show sample data structure
        st.subheader("📋 Expected Data Format")
        sample_data = {
            'Current Role': ['Developer', 'Analyst', 'Manager'],
            'Region': ['North', 'South', 'East'],
            'Avail Bucket': ['76-100%', '51-75%', '26-50%'],
            'Associate ID': ['EMP001', 'EMP002', 'EMP003'],
            'Associate Name': ['John Doe', 'Jane Smith', 'Bob Johnson'],
            'Current Availability': ['85%', '60%', '40%']
        }
        st.dataframe(pd.DataFrame(sample_data))

if __name__ == "__main__":
    main()