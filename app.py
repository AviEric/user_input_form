from flask import Flask, request, render_template
import openpyxl

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    name = request.form['name']
    age = request.form['age']
    enroll = request.form['enroll']

    # Load or create the workbook and select the active worksheet
    try:
        wb = openpyxl.load_workbook('user_data.xlsx')
    except FileNotFoundError:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["Name", "Age", "Enrollment Number"])
    else:
        ws = wb.active

    # Append the new data
    ws.append([name, age, enroll])

    # Save the workbook
    wb.save('user_data.xlsx')

    return 'Data saved successfully!'

if __name__ == '__main__':
    app.run(debug=True)
