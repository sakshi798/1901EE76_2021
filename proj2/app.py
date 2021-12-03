import os                                                                            #importing os module
from flask import Flask,render_template,redirect,request,flash                       #importing all the necessary modules from flask
from werkzeug.utils import secure_filename
from pathlib import Path                                                             #importing Path to get the path of home directory
from main import output,generate_pdf_range,generate_pdf_all                          #importing the important functions from main.py

app = Flask(__name__)                                                                #initialising the flask application
path = str(Path.home()) + "\\Desktop\\Proj2_Python\\sample_input"                    #defining the path where all the input and output
app.config['UPLOAD FOLDER'] = path                                                   #files will be saved on the user's desktop
app.secret_key = "CS384"                                                             #defining a secret key for the flask app

try:
    os.makedirs(app.config['UPLOAD FOLDER'])                                         #creating the Proj1_Python folder if it doesn't exist already
except:
    pass

@app.route("/",methods = ["GET","POST"])                                             #mapping the specific url with the associated function
def index():
    if request.method == "POST":

        req = request.form

        start = req.get("range_start")                                               #obtaining the roll no range input by the user in form
        end = req.get("range_end")
        start_roll = str(start)
        end_roll = str(end)

        f1 = request.files['csvfile1']                                               #obtaining the files input by the user in the form
        f2 = request.files['csvfile2']
        f3 = request.files['csvfile3']
        f4 = request.files['imagefile1']
        f5 = request.files['imagefile2']

        filename1 = secure_filename(f1.filename)
        filename2 = secure_filename(f2.filename)
        filename3 = secure_filename(f3.filename)
        filename4 = secure_filename(f4.filename)
        filename5 = secure_filename(f5.filename)

        path1 = os.path.join(app.config['UPLOAD FOLDER'],filename1)                 #creating a valid path for the input files inside the Proj2_Python folder   
        path2 = os.path.join(app.config['UPLOAD FOLDER'],filename2)
        path3 = os.path.join(app.config['UPLOAD FOLDER'],filename3)
        path4 = os.path.join(app.config['UPLOAD FOLDER'],filename4)
        path5 = os.path.join(app.config['UPLOAD FOLDER'],filename5)

        path1 = path1.replace('/','\\')                                             #replacing / by \\ to obtain a valid windows path
        path2 = path2.replace('/','\\')
        path3 = path3.replace('/','\\')
        path4 = path4.replace('/','\\')
        path5 = path5.replace('/','\\')

        if not os.path.exists(path1):                                               #saving the files to the created valid paths if not saved already
            f1.save(path1)

        if not os.path.exists(path2):
            f2.save(path2)

        if not os.path.exists(path3):
            f3.save(path3)

        if not os.path.exists(path4):
            f4.save(path4)

        if not os.path.exists(path5):
            f5.save(path5)

        output(path1,path2,path3)                                                   #calling the function to create output files to be used for generating transcripts
        
        if request.form['action'] == "Generate Transcripts for Given Roll Numbers": #calling the required function when the user clicks on a button
            list = generate_pdf_range(start_roll,end_roll,path4,path5)

        elif request.form['action'] == "Generate Transcripts for all students":
            generate_pdf_all(path4,path5)

        if len(list) != 0:                                                          #checking if the size of list is non-zero
            flash(list)                                                             #that is if all roll nos exist

        return redirect(request.url)

    return render_template('index.html')                                            #rendering the HTML that will display in the browser on the specific url

if __name__ == '__main__':
    app.run(debug=True)