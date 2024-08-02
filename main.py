from flask import Flask, request, send_file, render_template
import pandas as pd
import json
from io import BytesIO
from flask_cors import CORS
app = Flask(__name__)

CORS(app)

@app.route('/')
def upload_form():
    return render_template('home.html')

@app.route('/generate_excel', methods=['POST'])
def generate_excel():
    data = request.get_json()
    
    # Process the JSON data
    services = data.get('Groups', {}).get('Services', [])
    
    # Prepare data for Excel
    excel_data = {
        'No': [],
        'Service': [],
        'Quantity': [],
        'Price($)': [],
        'Detail': []
    }
    
    for i, service in enumerate(services, start=1):
        service_name = service['Service Name']
        description = service.get('Description', '')
        
        if description:
            service_display = f"{service_name} ({description})"
        else:
            service_display = service_name
        
        excel_data['No'].append(i)
        excel_data['Service'].append(service_display)
        excel_data['Quantity'].append(service['Properties'].get('Number of instances', 1))
        excel_data['Price($)'].append(round(float(service['Service Cost']['monthly']), 2))
        detail = ", ".join([f"{k}: {v}" for k, v in service['Properties'].items()])
        excel_data['Detail'].append(detail)
    
    # Create a DataFrame
    df = pd.DataFrame(excel_data)
    
    # Calculate the total price
    total_price = round(df['Price($)'].sum(), 2)
    total_row = ['Total', '', '', total_price, '']
    
    # Calculate the buffer price
    buffer_price = round(total_price + total_price * 0.10, 2)
    buffer_row = ['Buffer(10%)', '', '', buffer_price, '']
    
    # Append the total and buffer rows
    df.loc[len(df)] = total_row
    df.loc[len(df)] = buffer_row
    
    # Save to Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='EST', index=False)
        
        # Access the workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['EST']
        
        # Define a format for Arial font with size 14
        arial_format = workbook.add_format({
            'font_name': 'Arial',
            'font_size': 14,
            'border': 1,
            'align': 'center',
            'valign': 'vcenter'
        })
        
        # Define a format for the header
        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'top',
            'fg_color': '#FFA500',  # Orange color
            'border': 1,
            'align': 'center',  # Center the text
            'valign': 'vcenter',
            'font_name': 'Arial',
            'font_size': 14
        })
        
        # Define a format for the "Total" and "Buffer" rows in orange
        total_buffer_format = workbook.add_format({
            'bold': True,
            'bg_color': '#FFA500',  # Orange background
            'font_color': 'white',  # White text
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Arial',
            'font_size': 14
        })
        
        # Apply the format to the header row
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Define a format for the "Details" column with text wrapping
        wrap_format = workbook.add_format({
            'border': 1,
            'align': 'center',
            'valign': 'vcenter',
            'font_name': 'Arial',
            'font_size': 14,
            'text_wrap': True
        })
        
        # Apply the formats to the entire table
        for row in range(len(df) + 1):  # Include header row
            for col in range(len(df.columns)):
                if row == 0:
                    worksheet.write(row, col, df.columns[col], header_format)
                elif row == len(df) - 1 or row == len(df):  # "Total" and "Buffer" rows
                    worksheet.write(row, col, df.iloc[row-1, col], total_buffer_format)
                else:
                    if col == len(df.columns) - 1:  # "Detail" column
                        worksheet.write(row, col, df.iloc[row-1, col], wrap_format)
                    else:
                        worksheet.write(row, col, df.iloc[row-1, col], arial_format)
        
        # Adjust the column width
        worksheet.set_column('A:E', 20)
    
    output.seek(0)
    return send_file(output, download_name='aws_pricing.xlsx', as_attachment=True)




if __name__ == '__main__':
    app.run(debug=True)
