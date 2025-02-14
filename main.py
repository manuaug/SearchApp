#Import the required libraries

import pandas as pd
from flask import Flask,request,render_template

xl_file = pd.read_excel('servers.xlsx',sheet_name=None)

def search_data(search_term):
    results = pd.concat([xl_file[sheet].assign(sheet_name=sheet) for sheet in xl_file.keys()])
    results = results[results.apply(lambda row: row.astype(str).str.contains(search_term,case=False).any(),axis=1)]
    return results.to_dict(orient='records')

# Initialize flask app

app = Flask(__name__)

# Define routes

@app.route('/')

def search_page():
    return render_template('search.html')

@app.route('/results',methods=['GET','POST'])
def results():
    search_term = request.form['search_term']
    results = search_data(search_term)
    #Print the results
    return render_template('results.html',results=results)

# Run the app

if __name__ == '__main__':
    app.run(host='0.0.0.0',debug=True)
