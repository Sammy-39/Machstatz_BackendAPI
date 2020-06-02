from flask import Flask,request,jsonify,send_file
import requests,xlsxwriter

app = Flask(__name__)


@app.route('/total', methods=['GET'])
def total():
    source = requests.get('https://assignment-machstatz.herokuapp.com/excel').json()
    
    length,quantity,weight = {},{},{}
    
    day = str(request.args.get('day'))

    if day[:4] != '2020':
        day = day[6:] +'-'+ day[3:5] +'-'+ day[:2]

    length[day],quantity[day],weight[day] = 0.0,0.0,0.0

    for entry in source:
        day_s = str(dict(entry)['DateTime'][:10])
        if day_s == day:
            length[day]+=dict(entry)['Length']
            quantity[day]+=dict(entry)['Quantity']
            weight[day]+=dict(entry)['Weight']

    result = {'totalWeight': weight[day], 
              'totalLength': length[day],
              'totalQuantity': quantity[day] }

    return jsonify(result)

@app.route('/excelreport', methods=['GET'])
def get_excelreport():
    source_j = requests.get('https://assignment-machstatz.herokuapp.com/excel').json()
    
    workbook = xlsxwriter.Workbook('excelreport.xlsx')
    day = '00-00-0000'
    for entry in source_j:
        day_s = str(dict(entry)['DateTime'][:10])
        if day_s != day:
            worksheet = workbook.add_worksheet(day_s)
            worksheet.write(0, 0, 'Datetime')
            worksheet.write(0, 1, 'Length')
            worksheet.write(0, 2, 'Weight')
            worksheet.write(0, 3, 'Quantity')
            row_num = 1
            day = day_s
        worksheet.write(row_num, 0, dict(entry)['DateTime'])
        worksheet.write(row_num, 1, dict(entry)['Length'])
        worksheet.write(row_num, 2, dict(entry)['Weight'])
        worksheet.write(row_num, 3, dict(entry)['Quantity'])
        row_num+=1
    
    workbook.close()

    return send_file('excelreport.xlsx',mimetype='application/vnd.ms-excel')

if __name__ == '__main__':
   app.run(debug = True)