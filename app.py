import streamlit as st
import pandas as pd
import json
from io import BytesIO

def generate_pms_visualization(file_path=None, dataframe=None):
    """
    Generate an interactive HTML visualization from Excel data with bucket-based organization.

    Args:
        file_path: Path to the Excel file (optional if dataframe is provided)
        dataframe: Pre-loaded pandas DataFrame (optional if file_path is provided)

    Returns:
        HTML string of the visualization
    """
    # Load data
    if dataframe is not None:
        df = dataframe
    elif file_path is not None:
        df = pd.read_excel(file_path, engine='openpyxl')
    else:
        raise ValueError("Either file_path or dataframe must be provided")

    # Clean and process data
    df.columns = df.columns.str.strip()
    if 'Current Availability' in df.columns:
        df['Current Availability'] = df['Current Availability'].apply(
            lambda x: int(float(str(x).replace('%', ''))) if pd.notnull(x) else 0
        )
        df = df[df['Current Availability'] > 0]

    if df.empty:
        return """
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <title>PMS Resource Visualization</title>
            <style>
                body { font-family: 'Segoe UI', sans-serif; margin: 20px; }
                .empty-message {
                    text-align: center;
                    padding: 50px;
                    background: #f8f9fa;
                    border-radius: 8px;
                    color: #5f6368;
                    font-size: 18px;
                }
            </style>
        </head>
        <body>
            <div class="empty-message">No resources with availability greater than 0% found in the data.</div>
        </body>
        </html>
        """

    # Calculate total statistics
    total_associates = len(df)
    total_avg_availability = round(df['Current Availability'].mean(), 1)

    # Role mapping for standardization
    def map_role_name(role):
        role_str = str(role).upper().strip()
        if 'SCRUM' in role_str:
            return 'SCRUM'
        elif 'TPDL' in role_str:
            return 'TPDL'
        elif 'PGM' in role_str:
            return 'PGM'
        elif 'PM' in role_str:
            return 'PM'
        else:
            return str(role).strip()

    df['Mapped_Role'] = df['Current Role'].apply(map_role_name)

    # Get unique regions and roles
    all_regions = sorted(df['Region'].unique())
    all_roles = sorted(df['Mapped_Role'].unique())

    # Ensure standard roles are in a specific order
    standard_roles = ['Total', 'PGM', 'PM', 'SCRUM', 'TPDL']
    ordered_roles = [role for role in standard_roles if role in all_roles or role == 'Total']
    ordered_roles += [role for role in all_roles if role not in standard_roles]
    all_roles = ordered_roles

    # Create a 'Total' role category combining all data
    df_with_total = df.copy()

    # Define availability buckets
    all_buckets = ['76-100%', '51-75%', '26-50%', '0-25%']

    # Function to determine bucket
    def get_bucket(availability):
        if availability >= 76:
            return '76-100%'
        elif availability >= 51:
            return '51-75%'
        elif availability >= 26:
            return '26-50%'
        else:
            return '0-25%'

    df['Bucket'] = df['Current Availability'].apply(get_bucket)

    # Create data structure for the visualization
    dashboard_data = {
        'Total': {
            'count': total_associates,
            'avg_availability': total_avg_availability
        },
        'Regions': {},
        'Roles': {}
    }

    # Overall role statistics
    for role in all_roles:
        if role == 'Total':
            role_data = df
        else:
            role_data = df[df['Mapped_Role'] == role]

        if not role_data.empty:
            count = len(role_data)
            avg_avail = round(role_data['Current Availability'].mean(), 1)

            dashboard_data['Roles'][role] = {
                'count': count,
                'avg_availability': avg_avail,
                'buckets': {}
            }

            # Calculate bucket statistics for this role
            for bucket in all_buckets:
                bucket_data = role_data[role_data['Bucket'] == bucket]
                if not bucket_data.empty:
                    dashboard_data['Roles'][role]['buckets'][bucket] = {
                        'count': len(bucket_data),
                        'avg_availability': round(bucket_data['Current Availability'].mean(), 1),
                        'associates': bucket_data[['Associate ID', 'Associate Name', 'Current Availability', 'Region', 'Mapped_Role']].to_dict('records')
                    }
                else:
                    dashboard_data['Roles'][role]['buckets'][bucket] = {
                        'count': 0,
                        'avg_availability': 0,
                        'associates': []
                    }
        else:
            dashboard_data['Roles'][role] = {
                'count': 0,
                'avg_availability': 0,
                'buckets': {bucket: {'count': 0, 'avg_availability': 0, 'associates': []} for bucket in all_buckets}
            }

    # Region statistics
    for region in all_regions:
        region_data = df[df['Region'] == region]

        if not region_data.empty:
            count = len(region_data)
            avg_avail = round(region_data['Current Availability'].mean(), 1)

            dashboard_data['Regions'][region] = {
                'count': count,
                'avg_availability': avg_avail,
                'roles': {}
            }

            # Calculate role statistics for this region
            for role in all_roles:
                if role == 'Total':
                    role_region_data = region_data
                else:
                    role_region_data = region_data[region_data['Mapped_Role'] == role]

                if not role_region_data.empty:
                    dashboard_data['Regions'][region]['roles'][role] = {
                        'count': len(role_region_data),
                        'avg_availability': round(role_region_data['Current Availability'].mean(), 1),
                        'buckets': {}
                    }

                    # Calculate bucket statistics for this role in this region
                    for bucket in all_buckets:
                        bucket_data = role_region_data[role_region_data['Bucket'] == bucket]
                        if not bucket_data.empty:
                            dashboard_data['Regions'][region]['roles'][role]['buckets'][bucket] = {
                                'count': len(bucket_data),
                                'avg_availability': round(bucket_data['Current Availability'].mean(), 1),
                                'associates': bucket_data[['Associate ID', 'Associate Name', 'Current Availability', 'Mapped_Role']].to_dict('records')
                            }
                        else:
                            dashboard_data['Regions'][region]['roles'][role]['buckets'][bucket] = {
                                'count': 0,
                                'avg_availability': 0,
                                'associates': []
                            }
                else:
                    dashboard_data['Regions'][region]['roles'][role] = {
                        'count': 0,
                        'avg_availability': 0,
                        'buckets': {bucket: {'count': 0, 'avg_availability': 0, 'associates': []} for bucket in all_buckets}
                    }
        else:
            dashboard_data['Regions'][region] = {
                'count': 0,
                'avg_availability': 0,
                'roles': {role: {
                    'count': 0,
                    'avg_availability': 0,
                    'buckets': {bucket: {'count': 0, 'avg_availability': 0, 'associates': []} for bucket in all_buckets}
                } for role in all_roles}
            }

    # Generate HTML
    html = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>PMS Resource Visualization</title>
        <style>
            * { margin: 0; padding: 0; box-sizing: border-box; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
            body { background-color: #f0f2f5; color: #333; }
            .container { max-width: 1200px; margin: 0 auto; padding: 20px; }
            .header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; background-color: #fff; padding: 15px 20px; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
            .header-title { font-size: 20px; font-weight: 600; color: #333; }
            .total-box { display: flex; align-items: center; background-color: #e9ecef; border-radius: 6px; padding: 10px 15px; }
            .total-label { font-size: 14px; font-weight: 500; color: #555; }
            .total-value { font-size: 16px; font-weight: 600; color: #333; margin-left: 10px; }
            .tab-container { display: flex; border-bottom: 2px solid #e0e0e0; margin-bottom: 20px; }
            .tab { padding: 10px 20px; font-size: 15px; font-weight: 500; cursor: pointer; transition: all 0.2s; }
            .tab.active, .tab:hover { border-bottom: 2px solid #007bff; color: #007bff; }
            .tab-content { display: none; }
            .tab-content.active { display: block; }
            .role-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 20px; }
            .role-card { background-color: #fff; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); padding: 15px; cursor: pointer; transition: all 0.2s; }
            .role-card.active { border: 2px solid #007bff; }
            .role-card.no-data { background-color: #f8f9fa; cursor: not-allowed; }
            .role-close { font-size: 12px; font-weight: 500; color: #aaa; cursor: pointer; }
            .role-name { font-size: 16px; font-weight: 500; margin-bottom: 5px; }
            .role-stats { display: flex; justify-content: space-between; }
            .stat-label { font-size: 12px; color: #777; }
            .stat-number, .stat-percentage { font-size: 18px; font-weight: 600; color: #333; }
            .bucket-grid { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 20px; margin-top: 20px; }
            .bucket-card { background-color: #fff; border-radius: 8px; box-shadow: 0 2px 5px rgba(0,0,0,0.1); padding: 15px; cursor: pointer; transition: all 0.2s; }
            .bucket-card.bucket-76-100 { background-color: #ea4335; color: white; }
            .bucket-card.bucket-51-75 { background-color: #fbbc05; color: white; }
            .bucket-card.bucket-26-50 { background-color: #34a853; color: white; }
            .bucket-card.bucket-0-25 { background-color: #808080; color: white; }
            .bucket-name { font-size: 16px; font-weight: 500; margin-bottom: 5px; }
            .bucket-count, .bucket-avg { font-size: 14px; font-weight: 500; color: #fff; }
            .bucket-label { font-size: 12px; color: #ddd; }
            .no-data { text-align: center; color: #888; font-style: italic; margin: 20px 0; }
            .associates-modal { display: none; position: fixed; top: 50%; left: 50%; transform: translate(-50%, -50%); width: 80%; max-width: 1000px; background-color: #fff; border-radius: 8px; box-shadow: 0 5px 15px rgba(0,0,0,0.3); z-index: 1000; }
            .modal-header { display: flex; justify-content: space-between; align-items: center; padding: 15px 20px; border-bottom: 1px solid #e0e0e0; background-color: #f8f9fa; }
            .modal-title { font-size: 18px; font-weight: 600; color: #333; }
            .modal-subtitle { font-size: 14px; color: #777; margin-top: 5px; }
            .modal-close { font-size: 20px; font-weight: 500; cursor: pointer; }
            .modal-content { max-height: 70vh; overflow-y: auto; padding: 20px; }
            .associates-table { width: 100%; border-collapse: collapse; margin-top: 15px; }
            .associates-table th { background-color: #f1f3f4; color: #5f6368; text-align: left; padding: 12px 15px; font-weight: 500; border-bottom: 1px solid #e0e0e0; }
            .associates-table td { padding: 10px 15px; border-bottom: 1px solid #f0f0f0; }
            .associates-table tr:last-child td { border-bottom: none; }
            .associates-table tr:nth-child(even) { background-color: #f9f9f9; }
            .associates-table tr:hover { background-color: #f0f0f0; }
            .avail-indicator { display: inline-block; width: 12px; height: 12px; border-radius: 50%; margin-right: 5px; }
            .avail-76-100 { background-color: #ea4335; }
            .avail-51-75 { background-color: #fbbc05; }
            .avail-26-50 { background-color: #34a853; }
            .avail-0-25 { background-color: #808080; }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <div class="header-title">PMS Resource Dashboard</div>
                <div class="total-box">
                    <div class="total-label">Total:</div>
                    <div class="total-value">""" + f"{total_associates} Associates, {total_avg_availability}% Avg Availability" + """</div>
                </div>
            </div>
            <div class="tab-container">
                <div class="tab active" data-tab="overall">Overall</div>
                """ + ''.join([f'<div class="tab" data-tab="{region}">{region}</div>' for region in all_regions]) + """
            </div>
            <div id="overall-content" class="tab-content active">
                <div class="role-grid">
                    """ + ''.join([f"""
                    <div class="role-card {'no-data' if dashboard_data['Roles'][role]['count'] == 0 else ''}" data-role="{role}" onclick="showRoleBuckets('overall', '{role}')">
                        <div class="role-name">{role}</div>
                        <div class="role-stats">
                            <div>
                                <div class="stat-number">{dashboard_data['Roles'][role]['count']}</div>
                                <div class="stat-label">Associates</div>
                            </div>
                            <div>
                                <div class="stat-percentage">{dashboard_data['Roles'][role]['avg_availability']}%</div>
                                <div class="stat-label">Avg Availability</div>
                            </div>
                        </div>
                    </div>
                    """ for role in all_roles]) + """
                </div>
                <div id="bucket-container-overall"></div>
            </div>
            """ + ''.join([f"""
            <div id="{region}-content" class="tab-content">
                <div class="role-grid">
                    """ + ''.join([f"""
                    <div class="role-card {'no-data' if dashboard_data['Regions'][region]['roles'][role]['count'] == 0 else ''}" data-role="{role}" onclick="showRoleBuckets('{region}', '{role}')">
                        <div class="role-name">{role}</div>
                        <div class="role-stats">
                            <div>
                                <div class="stat-number">{dashboard_data['Regions'][region]['roles'][role]['count']}</div>
                                <div class="stat-label">Associates</div>
                            </div>
                            <div>
                                <div class="stat-percentage">{dashboard_data['Regions'][region]['roles'][role]['avg_availability']}%</div>
                                <div class="stat-label">Avg Availability</div>
                            </div>
                        </div>
                    </div>
                    """ for role in all_roles]) + """
                </div>
                <div id="bucket-container-""" + region + """"></div>
            </div>
            """ for region in all_regions]) + """
            <div id="associatesModal" class="associates-modal">
                <div class="modal-header">
                    <div class="modal-title" id="modalTitle"></div>
                    <div class="modal-subtitle" id="modalSubtitle"></div>
                    <div class="modal-close" onclick="closeModal()">&times;</div>
                </div>
                <div class="modal-content" id="modalBody"></div>
            </div>
        </div>
        <script>
            // Store the dashboard data
            const dashboardData = """ + json.dumps(dashboard_data) + """;

            // Tab switching functionality
            document.querySelectorAll('.tab').forEach(tab => {
                tab.addEventListener('click', function() {
                    document.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
                    this.classList.add('active');
                    document.querySelectorAll('.tab-content').forEach(content => {
                        content.classList.remove('active');
                        content.style.display = 'none';
                    });
                    const tabName = this.getAttribute('data-tab');
                    const contentElement = document.getElementById(tabName + '-content');
                    contentElement.classList.add('active');
                    contentElement.style.display = 'block';
                    document.getElementById('bucket-container-' + tabName).innerHTML = '';
                });
            });

            // Role card click to show buckets
            function showRoleBuckets(region, role) {
                document.querySelectorAll('.role-card').forEach(card => card.classList.remove('active'));
                const roleCards = document.querySelectorAll(`#${region}-content .role-card[data-role="${role}"]`);
                roleCards.forEach(card => card.classList.add('active'));

                let roleData;
                if (region === 'overall') {
                    roleData = dashboardData.Roles[role];
                } else {
                    roleData = dashboardData.Regions[region]?.roles[role];
                }

                if (!roleData || roleData.count === 0) {
                    document.getElementById(`bucket-container-${region}`).innerHTML =
                        '<div class="no-data">No data available for this selection</div>';
                    return;
                }

                let bucketsHTML = '<div class="bucket-grid">';
                const bucketOrder = ['76-100%', '51-75%', '26-50%', '0-25%'];

                for (const bucketName of bucketOrder) {
                    const bucketData = roleData.buckets[bucketName];
                    if (bucketData && bucketData.count > 0) {
                        let bucketClass = '';
                        if (bucketName === '76-100%') bucketClass = 'bucket-76-100';
                        else if (bucketName === '51-75%') bucketClass = 'bucket-51-75';
                        else if (bucketName === '26-50%') bucketClass = 'bucket-26-50';
                        else if (bucketName === '0-25%') bucketClass = 'bucket-0-25';

                        bucketsHTML += `
                            <div class="bucket-card ${bucketClass}" onclick="showAssociates('${region}', '${role}', '${bucketName}')">
                                <div class="bucket-name">${bucketName}</div>
                                <div class="bucket-count">${bucketData.count}</div>
                                <div class="bucket-label">Associates</div>
                                <div class="bucket-avg">${bucketData.avg_availability}%</div>
                                <div class="bucket-label">Avg Availability</div>
                            </div>
                        `;
                    }
                }

                bucketsHTML += '</div>';
                document.getElementById(`bucket-container-${region}`).innerHTML = bucketsHTML;
            }

            // Show associates in modal
            function showAssociates(region, role, bucket) {
                const modal = document.getElementById('associatesModal');
                const modalTitle = document.getElementById('modalTitle');
                const modalSubtitle = document.getElementById('modalSubtitle');
                const modalBody = document.getElementById('modalBody');

                let associates = [];
                let bucketData;

                if (region === 'overall') {
                    bucketData = dashboardData.Roles[role]?.buckets[bucket];
                } else {
                    bucketData = dashboardData.Regions[region]?.roles[role]?.buckets[bucket];
                }

                associates = bucketData?.associates || [];

                modalTitle.textContent = `${region === 'overall' ? 'Overall' : region} - ${role} - ${bucket}`;
                modalSubtitle.textContent = `${associates.length} associates, ${bucketData?.avg_availability || 0}% avg availability`;

                if (associates.length === 0) {
                    modalBody.innerHTML = '<div class="no-data">No associates found</div>';
                } else {
                    let tableHTML = `
                        <table class="associates-table">
                            <thead>
                                <tr>
                                    <th>Associate ID</th>
                                    <th>Associate Name</th>
                                    <th>Availability</th>`;

                    if (region === 'overall') {
                        tableHTML += '<th>Region</th>';
                    }

                    if (role === 'Total') {
                        tableHTML += '<th>Role</th>';
                    }

                    tableHTML += `
                                </tr>
                            </thead>
                            <tbody>`;

                    associates.sort((a, b) => b['Current Availability'] - a['Current Availability']);

                    associates.forEach(associate => {
                        const availability = associate['Current Availability'];
                        let availClass = '';

                        if (availability >= 76) availClass = 'avail-76-100';
                        else if (availability >= 51) availClass = 'avail-51-75';
                        else if (availability >= 26) availClass = 'avail-26-50';
                        else availClass = 'avail-0-25';

                        tableHTML += `
                            <tr>
                                <td>${associate['Associate ID'] || 'N/A'}</td>
                                <td>${associate['Associate Name'] || 'N/A'}</td>
                                <td>
                                    <span class="avail-indicator ${availClass}"></span>
                                    ${availability}%
                                </td>`;

                        if (region === 'overall') {
                            tableHTML += `<td>${associate['Region'] || 'N/A'}</td>`;
                        }

                        if (role === 'Total') {
                            tableHTML += `<td>${associate['Mapped_Role'] || 'N/A'}</td>`;
                        }

                        tableHTML += '</tr>';
                    });

                    tableHTML += '</tbody></table>';
                    modalBody.innerHTML = tableHTML;
                }

                modal.style.display = 'block';
            }

            // Close modal
            function closeModal() {
                document.getElementById('associatesModal').style.display = 'none';
            }

            // Close modal when clicking outside
            window.onclick = function(event) {
                const modal = document.getElementById('associatesModal');
                if (event.target === modal) {
                    modal.style.display = 'none';
                }
            };

            // Initialize
            document.addEventListener('DOMContentLoaded', function() {
                const overallTab = document.querySelector('.tab[data-tab="overall"]');
                overallTab.click();
                showRoleBuckets('overall', 'Total');
            });
        </script>
    </body>
    </html>
    """

    return html

def create_sample_data():
    """Create sample data for demonstration"""
    sample_data = {
        'Associate ID': [f'EMP{str(i).zfill(3)}' for i in range(1, 51)],
        'Associate Name': [f'Associate {i}' for i in range(1, 51)],
        'Current Role': ['Developer', 'Analyst', 'Manager', 'SCRUM Master', 'TPDL', 'PM', 'PGM'] * 7 + ['Developer'] * 2,
        'Region': ['North', 'South', 'East', 'West'] * 12 + ['North', 'South'],
        'Current Availability': [f'{i}%' for i in [85, 60, 40, 90, 75, 30, 95, 55, 25, 80] * 5]
    }
    return pd.DataFrame(sample_data)

def main():
    st.set_page_config(
        page_title="PMS Resource Analytics",
        page_icon="📊",
        layout="wide"
    )

    # Custom CSS for better styling
    st.markdown("""
    <style>
    .main-header {
        background: linear-gradient(90deg, #667eea 0%, #764ba2 100%);
        padding: 2rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-card {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        border-left: 4px solid #667eea;
    }
    .upload-section {
        background: #f8f9fa;
        padding: 2rem;
        border-radius: 10px;
        margin-bottom: 2rem;
    }
    </style>
    """, unsafe_allow_html=True)

    # Main header
    st.markdown("""
    <div class="main-header">
        <h1>🚀 Enterprise Resource Analytics Platform</h1>
        <p>Transform Excel data into interactive visualizations</p>
    </div>
    """, unsafe_allow_html=True)

    # Sidebar for file upload
    st.sidebar.header("📁 Data Source")

    # Option to choose data source
    data_source = st.sidebar.radio(
        "Choose data source:",
        ["Upload Excel File", "Use Sample Data"],
        help="Select how you want to provide the data"
    )

    uploaded_file = None
    df = None

    if data_source == "Upload Excel File":
        uploaded_file = st.sidebar.file_uploader(
            "Choose an Excel file",
            type=['xlsx', 'xls'],
            help="Upload your PMS data file with required columns"
        )
    else:
        if st.sidebar.button("🎯 Generate Sample Data", type="primary"):
            df = create_sample_data()
            st.sidebar.success("✅ Sample data generated!")

    # Main content area
    if uploaded_file is not None or df is not None:
        try:
            # Read the Excel file if uploaded
            if uploaded_file is not None:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
                st.success("✅ File uploaded successfully!")

            # Data validation
            required_columns = ['Current Role', 'Region', 'Associate ID', 'Associate Name', 'Current Availability']
            missing_columns = [col for col in required_columns if col not in df.columns]

            if missing_columns:
                st.error(f"❌ Missing required columns: {', '.join(missing_columns)}")
                st.info("Please ensure your Excel file contains all required columns.")
                return

            # Display data preview
            st.subheader("📋 Data Preview")
            with st.expander("Click to view data preview", expanded=True):
                st.dataframe(df.head(10), use_container_width=True)

            # Show data statistics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("📊 Total Records", len(df))
            with col2:
                st.metric("👥 Unique Roles", df['Current Role'].nunique())
            with col3:
                st.metric("🌍 Regions", df['Region'].nunique())
            with col4:
                if 'Current Availability' in df.columns:
                    avg_availability = df['Current Availability'].apply(
                        lambda x: float(str(x).replace('%', '')) if pd.notnull(x) else 0
                    ).mean()
                    st.metric("📈 Avg Availability", f"{avg_availability:.1f}%")

            # Generate visualization
            col1, col2, col3 = st.columns([1, 2, 1])
            with col2:
                if st.button("🎯 Generate Interactive Dashboard", type="primary", use_container_width=True):
                    with st.spinner("🔄 Generating interactive visualization..."):
                        # Generate the HTML visualization
                        html_content = generate_pms_visualization(dataframe=df)
                        # Render the HTML in Streamlit
                        st.components.v1.html(html_content, height=800, scrolling=True)
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")

if __name__ == "__main__":
    main()
