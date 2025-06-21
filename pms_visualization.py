# from allocation_script import data_use
import pandas as pd
import os
import json
import re

def generate_pms_visualization(file_path=None, dataframe=None):
    # """
    # Generate an interactive HTML visualization from Excel data.
    
    # Args:
    #     file_path: Path to the Excel file (optional if dataframe is provided)
    #     dataframe: Pre-loaded pandas DataFrame (optional if file_path is provided)
    
    # Returns:
    #     HTML string of the visualization
    # """
    # Load data either from file or use provided dataframe
    if dataframe is not None:
        df = dataframe
    elif file_path is not None:
        df = pd.read_excel(file_path, engine='openpyxl')
    else:
        raise ValueError("Either file_path or dataframe must be provided")
    
    # Clean column names to handle any whitespace issues
    df.columns = df.columns.str.strip()
    
    # Filter out associates with 0% availability first
    if 'Current Availability' in df.columns:
        df['Current Availability'] = df['Current Availability'].apply(
            lambda x: int(float(str(x).replace('%', ''))) if pd.notnull(x) else 0
        )
        # Remove rows with 0 availability
        df = df[df['Current Availability'] > 0]
        
    # If no data after filtering, return early with empty visualization
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
    
    # Get all possible values for each category to ensure we include empty buckets
    all_roles = sorted(df['Current Role'].unique())
    all_regions = sorted(df['Region'].unique())
    all_buckets = ['76-100%', '51-75%', '26-50%', '0-25%']  # Fixed order for buckets
    
    # Create hierarchical structure
    data_dict = {
        'name': 'PMS',
        'children': {}
    }
    
    # Create a dynamic structure based on actual data rather than all combinations
    # We'll populate it as we process the data
    for role in all_roles:
        if role not in data_dict['children']:
            data_dict['children'][role] = {'name': role, 'children': {}}
            
        for region in all_regions:
            if region not in data_dict['children'][role]['children']:
                data_dict['children'][role]['children'][region] = {'name': region, 'children': {}}
                
            # We don't pre-create buckets - they'll be added only when needed
    
    # Now populate with actual data
    for _, row in df.iterrows():
        try:
            role = str(row['Current Role']).strip()
            region = str(row['Region']).strip()
            
            # Handle availability and bucket
            availability = 0
            if 'Current Availability' in row and pd.notnull(row['Current Availability']):
                availability = int(float(str(row['Current Availability']).replace('%', '')))
            
            # Determine which bucket the availability falls into
            bucket = ''
            if availability >= 76:
                bucket = '76-100%'
            elif availability >= 51:
                bucket = '51-75%'
            elif availability >= 26:
                bucket = '26-50%'
            else:
                bucket = '0-25%'
                
            # Extract associate details (safely handling potential missing values)
            associate_id = str(row.get('Associate ID', '')).strip() if pd.notnull(row.get('Associate ID', '')) else ''
            associate_name = str(row.get('Associate Name', '')).strip() if pd.notnull(row.get('Associate Name', '')) else ''
            
            # Skip associates with 0% availability
            if availability == 0:
                continue
                
            # Add associate details as an object for the table
            associate_info = {
                'id': associate_id,
                'name': associate_name,
                'availability': availability
            }
            
            # Only add associate if we have valid role, region and bucket
            if role and region and bucket:
                # Create the bucket if it doesn't exist
                if bucket not in data_dict['children'][role]['children'][region]['children']:
                    data_dict['children'][role]['children'][region]['children'][bucket] = {
                        'name': bucket,
                        'children': {},
                        'associates': []
                    }
                
                data_dict['children'][role]['children'][region]['children'][bucket]['associates'].append(associate_info)
                
        except Exception as e:
            print(f"Error processing row: {row}, Error: {e}")
    
    # Generate HTML
    html = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>PMS Resource Visualization</title>
        <style>
            body {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                margin: 0;
                padding: 20px;
                background-color: #f8f9fa;
                color: #212529;
            }
            
            .container {
                max-width: 1600px;
                margin: 0 auto;
            }
            
            .header {
                background: linear-gradient(135deg, #2b5876 0%, #4e4376 100%);
                color: white;
                padding: 20px;
                border-radius: 8px;
                margin-bottom: 25px;
                box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
            }
            
            h1 {
                margin: 0;
                font-size: 24px;
            }
            
            .info {
                margin-top: 5px;
                opacity: 0.8;
                font-size: 14px;
            }
            
            /* Vertical Tree Styling */
            .org-tree {
                display: flex;
                justify-content: flex-start;
                overflow-x: auto;
                padding: 20px 0;
            }
            
            .org-tree ul {
                padding-left: 20px;
                position: relative;
            }
            
            .org-tree li {
                list-style-type: none;
                position: relative;
                padding: 10px 0 0 10px;
            }
            
            .org-tree li::before {
                content: "";
                position: absolute;
                top: 0;
                left: 0;
                border-left: 1px solid #ccc;
                height: 100%;
                width: 1px;
            }
            
            .org-tree li:last-child::before {
                height: 20px;
            }
            
            .org-tree li::after {
                content: "";
                position: absolute;
                top: 20px;
                left: 0;
                border-top: 1px solid #ccc;
                width: 10px;
            }
            
            .org-tree > ul > li::before,
            .org-tree > ul > li::after {
                border: none;
            }
            
            .node {
                display: inline-block;
                padding: 5px 10px;
                border-radius: 5px;
                font-weight: 500;
                position: relative;
                cursor: pointer;
                transition: all 0.2s;
                min-width: 120px; /* Standardized width for better alignment */
                text-align: center; /* Center text for consistency */
            }
            
            .node:hover {
                box-shadow: 0 0 5px rgba(0, 0, 0, 0.2);
            }
            
            /* Node Styling by Level */
            .node-root {
                background: linear-gradient(135deg, #2b5876 0%, #4e4376 100%);
                color: white;
                font-weight: bold;
                padding: 8px 16px;
                border-radius: 20px;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
                min-width: 150px;
            }
            
            .node-role {
                background: linear-gradient(135deg, #1a73e8 0%, #1559b7 100%);
                color: white;
                border-radius: 6px;
                min-width: 150px; /* Wider to accommodate longer role names */
                text-align: center;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
                padding: 8px 12px;
            }
            
            .node-region {
                background: linear-gradient(135deg, #666699 0%, #5c5c8a 100%);
                color: white;
                border-radius: 6px;
                min-width: 120px;
                text-align: center;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            }
            
            .node-bucket {
                color: white;
                border-radius: 6px;
                min-width: 120px;
                text-align: center;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            }
            
            /* Bucket colors by availability */
            .bucket-high {
                background: linear-gradient(135deg, #ea4335 0%, #e72918 100%);
            }

            .bucket-medium {
                background: linear-gradient(135deg, #fbbc05 0%, #e2aa03 100%);
            }

            .bucket-low {
                background: linear-gradient(135deg, #70db70 0%, #5cd65c 100%);
            }

            .bucket-very-low {
                background: linear-gradient(135deg, #34a853  0%, #2a8943 100%);
            }

            
            /* Table Styling */
            .associates-table {
                border-collapse: collapse;
                width: 100%;
                margin-top: 15px;
                background-color: white;
                border-radius: 5px;
                box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
                overflow: hidden;
            }
            
            .associates-table th {
                background-color: #f1f3f4;
                color: #5f6368;
                text-align: left;
                padding: 12px 15px;
                font-weight: 500;
                border-bottom: 1px solid #e0e0e0;
                white-space: nowrap;
            }
            
            .associates-table td {
                padding: 10px 15px;
                border-bottom: 1px solid #f0f0f0;
            }
            
            .associates-table tr:last-child td {
                border-bottom: none;
            }
            
            .associates-table tr:nth-child(even) {
                background-color: #f9f9f9;
            }
            
            .associates-table tr:hover {
                background-color: #f5f5f5;
            }
            
            /* Toggle functionality */
            .nested {
                display: none;
                margin-top: 10px;
                margin-left: 5px;
                padding: 10px;
                background-color: rgba(255, 255, 255, 0.7);
                border-radius: 8px;
                box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
                min-width: 300px;
            }
            
            .active {
                display: block;
            }
            
            /* Availability indicator in table */
            .availability-indicator {
                display: inline-block;
                width: 12px;
                height: 12px;
                border-radius: 50%;
                margin-right: 8px;
                vertical-align: middle;
            }
            
            .avail-high {
                background-color: #ea4335; /* Red */
            }

            .avail-medium {
                background-color: #fbbc05; /* Amber */
            }

            .avail-low {
                background-color: #70db70; /* Light Green */
            }

            .avail-very-low {
                background-color: #34a853; /* Dark Green */
            }

            
            /* Empty state styling */
            .empty-state {
                padding: 15px;
                text-align: center;
                color: #5f6368;
                font-style: italic;
                background-color: #f9f9f9;
                border-radius: 5px;
                margin-top: 10px;
            }
            
            /* Expand/collapse icons */
            .toggle-icon {
                margin-right: 8px;
                display: inline-block;
                width: 16px;
                text-align: center;
                font-weight: bold;
            }
            
            /* Responsive adjustments */
            @media (max-width: 768px) {
                .org-tree ul {
                    padding-left: 15px;
                }
                
                .node {
                    padding: 4px 8px;
                    font-size: 14px;
                }
            }
            
            /* Summary table at the top */
            .summary-container {
                margin-bottom: 30px;
                background-color: white;
                padding: 20px;
                border-radius: 8px;
                box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
            }
            
            .summary-table {
                width: 100%;
                border-collapse: collapse;
                margin-top: 15px;
            }
            
            .summary-table th {
                background-color: #f1f3f4;
                padding: 12px 15px;
                text-align: left;
                border-bottom: 2px solid #e0e0e0;
            }
            
            .summary-table td {
                padding: 10px 15px;
                border-bottom: 1px solid #f0f0f0;
            }
            
            .summary-title {
                font-size: 18px;
                font-weight: 500;
                color: #202124;
                margin-bottom: 10px;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="header">
                <h1>PMS Resource Visualization</h1>
                <div class="info">Click on nodes to expand/collapse the tree structure and view associate details</div>
            </div>
            
            <div class="org-tree">
    """
    
    # Build the tree structure recursively
    html += _build_tree_html(data_dict)
    
    html += """
            </div>
        </div>
        <script>
            document.addEventListener("DOMContentLoaded", function() {
                // Add event listeners to all nodes
                var nodes = document.querySelectorAll(".node");
                nodes.forEach(function(node) {
                    node.addEventListener("click", function(event) {
                        // Prevent event from bubbling up
                        event.stopPropagation();
                        
                        // Toggle visibility of child elements
                        var parent = this.parentElement;
                        var nested = parent.querySelector(".nested");
                        if (nested) {
                            nested.classList.toggle("active");
                            
                            // Toggle icon
                            var icon = this.querySelector(".toggle-icon");
                            if (icon) {
                                if (nested.classList.contains("active")) {
                                    icon.innerHTML = "−";
                                } else {
                                    icon.innerHTML = "+";
                                }
                            }
                        }
                    });
                });
                
                // Expand root node by default
                var rootNode = document.querySelector(".node-root");
                if (rootNode) {
                    rootNode.click();
                }
            });
        </script>
    </body>
    </html>
    """
    
    return html

def _build_tree_html(node, level=0):
    """Helper function to recursively build the HTML tree"""
    if level == 0:
        # Root node (PMS)
        html = '<ul><li><div class="node node-root"><span class="toggle-icon">+</span>PMS</div><ul class="nested">'
        
        # Sort the roles for consistent display
        sorted_roles = sorted(node['children'].keys())
        for role in sorted_roles:
            html += _build_tree_html(node['children'][role], level+1)
            
        html += '</ul></li></ul>'
    elif level == 1:
        # Role level
        html = f'<li><div class="node node-role"><span class="toggle-icon">+</span>{node["name"]}</div><ul class="nested">'
        
        # Sort the regions
        sorted_regions = sorted(node['children'].keys())
        for region in sorted_regions:
            html += _build_tree_html(node['children'][region], level+1)
            
        html += '</ul></li>'
    elif level == 2:
        # Region level
        html = f'<li><div class="node node-region"><span class="toggle-icon">+</span>{node["name"]}</div><ul class="nested">'
        
        # Sort buckets in specified order: '76-100%', '51-75%', '26-50%', '0-25%'
        def bucket_sort_key(bucket):
            if '76-100%' in bucket:
                return 0
            elif '51-75%' in bucket:
                return 1
            elif '26-50%' in bucket:
                return 2
            elif '0-25%' in bucket:
                return 3
            else:
                return 4
        
        # Only include buckets that have associates
        non_empty_buckets = []
        for bucket, bucket_data in node['children'].items():
            if bucket_data.get('associates', []):
                non_empty_buckets.append(bucket)
                
        sorted_buckets = sorted(non_empty_buckets, key=bucket_sort_key)
        
        # Check if there are any non-empty buckets
        if sorted_buckets:
            for bucket in sorted_buckets:
                html += _build_tree_html(node['children'][bucket], level+1)
        else:
            # Add an empty state message if no associates in any bucket
            html += '<li><div class="empty-state">No associates found in any availability bucket</div></li>'
            
        html += '</ul></li>'
    elif level == 3:
        # Bucket level with table of associates
        # Add appropriate CSS class based on bucket percentage
        bucket_class = "node-bucket "
        if '76-100%' in node['name']:
            bucket_class += "bucket-high"
        elif '51-75%' in node['name']:
            bucket_class += "bucket-medium"
        elif '26-50%' in node['name']:
            bucket_class += "bucket-low"
        elif '0-25%' in node['name']:
            bucket_class += "bucket-very-low"
        
        html = f'<li><div class="node {bucket_class}"><span class="toggle-icon">+</span>{node["name"]}</div>'
        
        # Add nested content with improved table
        html += '<div class="nested">'
        
        # Sort associates by availability (descending)
        sorted_associates = sorted(
            node.get('associates', []), 
            key=lambda x: (
                -float(x.get('availability', 0)),
                x.get('name', '')
            )
        )
        
        if sorted_associates:
            # Create an improved table for associates with proper formatting
            html += '''
            <table class="associates-table">
                <thead>
                    <tr>
                        <th>Associate ID</th>
                        <th>Associate Name</th>
                        <th>Current Availability</th>
                    </tr>
                </thead>
                <tbody>
            '''
            
            for associate in sorted_associates:
                # Determine availability class for the indicator
                avail_class = ""
                avail = float(associate.get('availability', 0))
                if avail >= 76:
                    avail_class = "avail-high"
                elif avail >= 51:
                    avail_class = "avail-medium"
                elif avail >= 26:
                    avail_class = "avail-low"
                else:
                    avail_class = "avail-very-low"
                
                # Generate table row with associate details
                html += f'''
                <tr>
                    <td>{associate.get('id', '')}</td>
                    <td>{associate.get('name', '')}</td>
                    <td><span class="availability-indicator {avail_class}"></span>{associate.get('availability', 0)}%</td>
                </tr>
                '''
            
            html += '''
                </tbody>
            </table>
            '''
        else:
            # Display empty state message
            html += '<div class="empty-state">No associates in this bucket</div>'
        
        html += '</div></li>'
    
    return html

def process_excel_file(file_path):
    """Process an Excel file and generate visualization"""
    html = generate_pms_visualization(file_path=file_path)
    
    # Save the HTML to a file
    output_file = os.path.splitext(file_path)[0] + '_ImprovedTree.html'
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html)
    
    print(f"Visualization saved to: {output_file}")
    return output_file

# If running as a script
if __name__ == "__main__":
    import sys
    
    if len(sys.argv) > 1:
        file_path = sys.argv[1]
        process_excel_file(file_path)
    else:
        print("Please provide the path to the Excel file as an argument.")
        print("Example: python pms_visualization.py path/to/excel_file.xlsx")