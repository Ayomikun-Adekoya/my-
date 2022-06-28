from flask import Flask, redirect, url_for, render_template, request
from werkzeug.utils import secure_filename

app = Flask(__name__)
@app.route("/")
def home():
    return render_template("front_end.html")

allowed_extensions = ['xlsx']

def check_file_extension(filename):
    return filename.split('.')[-1] in allowed_extensions



@app.route("/abs_value/")
def abs_value():
    return render_template("Abs_values.html")


@app.route("/protein_value/")
def protein_value():
    return render_template("protein_value.html")



@app.route('/upload', methods = ['GET', 'POST'])

def uploadfile():
   if request.method == 'POST': # check if the method is post
      f = request.files['file_1', 'file_2'] # get the file from the files object
      # Saving the file in the required destination
      if check_file_extension(f.filename):
         f.save(secure_filename(f.filename)) # this will secure the file
         return ('file uploaded successfully') # Display thsi message after uploading

      else:
         return 'The file extension is not allowed'






if __name__ == "__main__":
    app.run(debug=True)

