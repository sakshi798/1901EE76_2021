import os
import csv                                                                              #importing os and csv modules
from flask import Flask,request,render_template,redirect,flash                          #importing all the necessary modules from flask
from werkzeug.utils import secure_filename                                              
from main import create_concise_mk,create_rollnowise,send_email                         #importing the important functions from main.py
from pathlib import Path                                                                #importing Path to get the path of home directory

def valid_file(path1):                                                                  #defining a function to check whether master_roll
    with open(path1,'r') as f:                                                          #contains a row with roll no as ANSWER
        reader = csv.reader(f)
        for row in reader:
            if row[0] != 'roll':
                if row[0] == 'ANSWER':
                    return True

    return False

app = Flask(__name__)                                                                  #initialising the flask application
path = str(Path.home()) + "\\Desktop\\Proj1_Python\\sample_input"                      #defining the path where all the input and output
app.config['UPLOAD FOLDER'] = path                                                     #files will be saved on the user's desktop
app.secret_key = "CS384"                                                               #defining a secret key for the flask app

try:
    os.makedirs(app.config['UPLOAD FOLDER'])                                           #creating the Proj1_Python folder if it doesn't exist already
except:
    pass

@app.route('/',methods = ['GET','POST'])                                               #mapping the specific url with the associated function
def index():

    if request.method == "POST":

        req = request.form

        pos = req.get("correct")                                                    #obtaining the value of correct and wrong marks input by
        neg = req.get("wrong")                                                      #the user in the form

        f1 = request.files['csvfile1']                                              #obtaining the files input by the user in the form
        f2 = request.files['csvfile2']

        filename1 = secure_filename(f1.filename)                                    
        filename2 = secure_filename(f2.filename)

        path1 = os.path.join(app.config['UPLOAD FOLDER'],filename1)                #creating a valid path for the input files inside the Proj1_Python folder
        path2 = os.path.join(app.config['UPLOAD FOLDER'],filename2)
        path1 = path1.replace('/','\\')                                            #replacing / by \\ to obtain a valid windows path
        path2 = path2.replace('/','\\')

        if not os.path.exists(path1):
            f1.save(path1)                                                         #saving the files to the created valid paths if not saved already

        if not os.path.exists(path2):
            f2.save(path2)

        if not valid_file(path1):                                                #if the master_roll.csv file doesn't contain ANSWER then we flash an error 
            flash("no roll number with ANSWER is present, can't process!")       #message and refresh the page
            return redirect(request.url)

        if request.form['action'] == "Generate Roll number wise Marksheet":      #calling the required function when the user clicks on a button
            create_rollnowise(path1,path2,pos,neg)                               

        elif request.form['action'] == "Generate Concise Marksheet with Roll Num,Obtained Marks,marks after negative":
            create_concise_mk(path1,path2,pos,neg)

        elif request.form['action'] == "Send Email":
            send_email()

        else:
            redirect(request.url)

    return render_template('index.html')                                        #rendering the HTML that will display in the browser on the specific url

if __name__ == '__main__':
    app.run(debug=True)                                                         
