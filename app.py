from flask import Flask, request, send_file
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
import io
import os
from datetime import datetime

app = Flask(__name__)

# Path to your SF2 template (place in same folder as this file)
TEMPLATE_PATH = 'SF2_template.xlsx'

def get_month_index(month_name):
    """Convert month name to index"""
    months = ['January', 'February', 'March', 'April', 'May', 'June',
              'July', 'August', 'September', 'October', 'November', 'December']
    try:
        return months.index(month_name) + 1
    except ValueError:
        return 0

def column_letter(col_num):
    """Convert column number to letter (0=A, 1=B, etc)"""
    return chr(65 + col_num)

@app.route('/api/generate-sf2', methods=['POST'])
def generate_sf2():
    """Generate SF2 report from attendance data"""
    try:
        # Get data from request
        data = request.json
        month = data.get('month')
        year = data.get('year')
        students = data.get('students', [])
        
        print(f"Generating SF2 for {month} {year}")
        print(f"Processing {len(students)} students")
        
        # Validate month
        month_index = get_month_index(month)
        if month_index == 0:
            return {'error': f'Invalid month: {month}'}, 400
        
        # Check if template exists
        if not os.path.exists(TEMPLATE_PATH):
            return {'error': f'Template not found: {TEMPLATE_PATH}'}, 500
        
        # Load template
        wb = load_workbook(TEMPLATE_PATH)
        ws = wb.active
        
        print(f"Template loaded: {ws.title}")
        
        # Store merged cells to restore later
        merged_cells = list(ws.merged_cells.ranges)
        print(f"Found {len(merged_cells)} merged cell ranges")
        
        # Unmerge all cells temporarily (except header rows to preserve images)
        for merged_range in merged_cells:
            range_str = str(merged_range)
            # Only unmerge data rows (rows 14+), skip header rows
            if int(range_str.split(':')[0][1:]) >= 14 or int(range_str.split(':')[1][1:]) >= 14:
                try:
                    ws.unmerge_cells(range_str)
                except:
                    pass
        
        # Step 1: Write month/year to O6
        ws['O6'] = f"{month} {year}"
        print(f"✅ Written month/year to O6: {month} {year}")
        
        # Step 2: Get days in month
        if month_index == 12:
            next_month_date = f"{year + 1}-01-01"
        else:
            next_month_date = f"{year}-{month_index + 1:02d}-01"
        
        from datetime import datetime, timedelta
        last_day = datetime.strptime(next_month_date, "%Y-%m-%d") - timedelta(days=1)
        days_in_month = last_day.day
        
        print(f"Days in month: {days_in_month}")
        
        # Step 3: Write day numbers to row 10 (row index 9) starting from column D (3)
        for day in range(1, days_in_month + 1):
            col_index = 3 + (day - 1)  # Start from D (index 3)
            cell = ws.cell(row=10, column=col_index)
            cell.value = day
        print(f"✅ Written day numbers 1-{days_in_month} to row 10")
        
        # Step 4: Separate students by gender
        male_students = [s for s in students if s.get('gender', '').upper() == 'MALE']
        female_students = [s for s in students if s.get('gender', '').upper() == 'FEMALE']
        
        # Sort alphabetically
        male_students.sort(key=lambda s: s.get('name', ''))
        female_students.sort(key=lambda s: s.get('name', ''))
        
        print(f"Males: {len(male_students)}, Females: {len(female_students)}")
        
        # Step 5: Write MALE students (rows 14-34)
        male_start_row = 14
        for idx, student in enumerate(male_students[:21]):  # Max 21 male students
            row = male_start_row + idx
            
            # Write row number in column A
            ws.cell(row=row, column=1).value = idx + 1
            
            # Write student name in column B
            ws.cell(row=row, column=2).value = student.get('name', '')
            
            print(f"Male #{idx+1}: {student.get('name')} at row {row}")
            
            # Write attendance
            attendance = student.get('attendance', [])
            for att in attendance:
                try:
                    att_date = datetime.strptime(att.get('date'), '%Y-%m-%d')
                    
                    # Check if date is in current month/year
                    if att_date.month == month_index and att_date.year == year:
                        day = att_date.day
                        col_index = 3 + (day - 1)  # Column D onwards
                        cell = ws.cell(row=row, column=col_index)
                        
                        status = att.get('status', '')
                        if status == 'Absent':
                            cell.value = 'x'
                            cell.font = Font(bold=True, color='FF0000')  # Red bold x
                        elif status == 'Late':
                            cell.value = 'T'
                            apply_half_shading(cell, 'darkUp')  # Upper half shaded
                        elif status == 'Cutting Class':
                            cell.value = 'T'
                            apply_half_shading(cell, 'darkDown')  # Lower half shaded
                        # Present = leave blank
                except Exception as e:
                    print(f"Error processing attendance: {e}")
        
        # Step 6: Write FEMALE students (rows 36-60)
        female_start_row = 36
        for idx, student in enumerate(female_students[:25]):  # Max 25 female students
            row = female_start_row + idx
            
            # Write row number in column A
            ws.cell(row=row, column=1).value = idx + 1
            
            # Write student name in column B
            ws.cell(row=row, column=2).value = student.get('name', '')
            
            print(f"Female #{idx+1}: {student.get('name')} at row {row}")
            
            # Write attendance
            attendance = student.get('attendance', [])
            for att in attendance:
                try:
                    att_date = datetime.strptime(att.get('date'), '%Y-%m-%d')
                    
                    # Check if date is in current month/year
                    if att_date.month == month_index and att_date.year == year:
                        day = att_date.day
                        col_index = 3 + (day - 1)  # Column D onwards
                        cell = ws.cell(row=row, column=col_index)
                        
                        status = att.get('status', '')
                        if status == 'Absent':
                            cell.value = 'x'
                            cell.font = Font(bold=True, color='FF0000')  # Red bold x
                        elif status == 'Late':
                            cell.value = 'T'
                            apply_half_shading(cell, 'darkUp')  # Upper half shaded
                        elif status == 'Cutting Class':
                            cell.value = 'T'
                            apply_half_shading(cell, 'darkDown')  # Lower half shaded
                        # Present = leave blank
                except Exception as e:
                    print(f"Error processing attendance: {e}")
        
        # Step 7: Re-merge cells (only data rows that were unmerged)
        for merged_range in merged_cells:
            range_str = str(merged_range)
            # Only re-merge if we unmerged it
            if int(range_str.split(':')[0][1:]) >= 14 or int(range_str.split(':')[1][1:]) >= 14:
                try:
                    ws.merge_cells(range_str)
                except Exception as e:
                    print(f"Warning: Could not re-merge {range_str}: {e}")
        
        # Step 8: Save to memory
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        print(f"✅ SF2 report generated successfully")
        
        # Return file
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=f"SF2_{month}_{year}.xlsx"
        )
    
    except Exception as e:
        print(f"❌ Error: {str(e)}")
        return {'error': str(e)}, 500

@app.route('/health', methods=['GET'])
def health():
    """Health check endpoint"""
    return {'status': 'ok', 'timestamp': datetime.now().isoformat()}

if __name__ == '__main__':
    print("Starting SF2 Backend Server...")
    print(f"Template file: {TEMPLATE_PATH}")
    app.run(debug=True, host='0.0.0.0', port=5000)
