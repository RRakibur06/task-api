from flask import Flask
from flask_restful import  Api
from flask import Flask, redirect, url_for, request
from spire.presentation.common import *
from spire.presentation import *
import os
import xlsxwriter  
from ultralytics import YOLO

app = Flask(__name__)
api = Api(app)

@app.route('/api/input',methods = ['POST'])
def input():
    file = request.files["file"]
    file.save(file.filename)
    x = extract(file.filename)
    return "File submitted: " + file.filename

def extract(filename):
    ppt = Presentation()
    ppt.LoadFromFile(filename)

    if not os.path.isdir("./PPT_Image"):
        os.makedirs("./PPT_Image")

    for i, image in enumerate(ppt.Images):
        ImageName = "./PPT_Image/Images_"+str(i)+".png"
        image.Image.Save(ImageName)

    x = createExcel(filename)
    ppt.Dispose()

def createExcel(filename):
    workbook = xlsxwriter.Workbook("./"+filename+".xlsx")     
    sheet = workbook.add_worksheet()     
    model = YOLO("yolov8m.pt")
        
    sheet.write(0, 0, 'File Name')
    sheet.write(0, 1, 'Contents')   
    fileRow = 1    
    objRow = 1    
    path = os.path.dirname(os.path.abspath(__file__))
      
    joinedPath = os.path.join(path, 'PPT_Image')
    content = os.listdir(joinedPath)
    print(content)    
    for i in content :     
        sheet.write(fileRow, 0, i) 
        fileRow += 1  
        results = model.predict(os.path.join(joinedPath, i))
        result = results[0]
        print(result)
        detected_objects = []
        for box in result.boxes:
            detected_objects.append(result.names.get(box.cls[0].item()))
        print(detected_objects)
        col = 1
        for data in detected_objects:
            sheet.write(objRow, col, data)
            col += 1
        objRow += 1
                
    workbook.close()     


if __name__ == '__main__':
    app.run(debug=True)