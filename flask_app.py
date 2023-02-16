from flask import Flask, render_template, request, send_file
import os
from pptx import Presentation
import pandas as pd
from PIL import Image
import io

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['STATIC_FOLDER'] = 'static'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download_template')
def download_template():
    return send_file(os.path.join(app.config['STATIC_FOLDER'], 'template.xlsx'), as_attachment=True)

@app.route('/upload_files', methods=['POST'])
def upload_files():
    # Handle uploaded data.xlsx file
    data_file = request.files['data_file']
    data_path = os.path.join(app.config['UPLOAD_FOLDER'], 'data.xlsx')
    data_file.save(data_path)

    # Handle uploaded picture files
    picture_paths = []
    for picture_file in request.files.getlist('picture_files'):
        picture_path = os.path.join(app.config['UPLOAD_FOLDER'], picture_file.filename)
        picture_file.save(picture_path)
        picture_paths.append(picture_path)

    # Load data from Excel file
    data = pd.read_excel(data_path)
    
    # Load PowerPoint template and select slide layout
    prs = Presentation(os.path.join(app.config['STATIC_FOLDER'], 'template.pptx'))
    layout = prs.slide_layouts[0]

    # Loop over rows in Excel file and create new slide for each row
    for index, row in data.iterrows():
        slide = prs.slides.add_slide(layout)

        # Populate text placeholders on slide with values from Excel row
        slide.shapes.placeholders[14].text = str(row['Titel'])
        slide.shapes.placeholders[10].text = str(row['Site'])
        slide.shapes.placeholders[11].text = str(row['Building'])
        slide.shapes.placeholders[12].text = str(row['Office'])
        slide.shapes.placeholders[13].text = str(row['Mezzanine'])
        slide.shapes.placeholders[17].text = str(row['Parking'])
        slide.shapes.placeholders[18].text = str(row['Environment category'])
        slide.shapes.placeholders[19].text = str(row['Maximum building height'])
        slide.shapes.placeholders[20].text = str(row['Clear height'])
        slide.shapes.placeholders[21].text = str(row['Floor load'])
        slide.shapes.placeholders[22].text = str(row['Floor flatness'])
        slide.shapes.placeholders[23].text = str(row['Loading docks'])
        slide.shapes.placeholders[24].text = str(row['Overhead doors'])
        slide.shapes.placeholders[25].text = str(row['Sprinkler'])

        # Calculate the index of the first picture on this slide
        first_picture_index = 15
        
        # Populate the first picture placeholder on slide with image1 from folder
    pictures_dir = app.config['UPLOAD_FOLDER']
    image1_filename = f'image{index*2+1}'
    image1_path = None

    for filename in os.listdir(pictures_dir):
        if filename.startswith(image1_filename):
            image1_path = os.path.join(pictures_dir, filename)
            break

    if image1_path is None:
        print(f'Error: Could not find {image1_filename} in {pictures_dir}')
    else:
        with open(image1_path, 'rb') as f:
            img_bytes = f.read()
            image1_ext = os.path.splitext(image1_path)[1].lower()

            if image1_ext == '.webp':
                img = Image.open(io.BytesIO(img_bytes)).convert('RGB')
                img_bytes = io.BytesIO()
                img.save(img_bytes, format='JPEG')
                img_bytes = img_bytes.getvalue()

            slide.shapes.placeholders[first_picture_index].insert_picture(io.BytesIO(img_bytes))

    # Populate the second picture placeholder on slide with image2 from folder
    image2_filename = f'image{index*2+2}'
    image2_path = None

    for filename in os.listdir(pictures_dir):
        if filename.startswith(image2_filename):
            image2_path = os.path.join(pictures_dir, filename)
            break

    if image2_path is None:
        print(f'Error: Could not find {image2_filename} in {pictures_dir}')
    else:
        with open(image2_path, 'rb') as f:
            img_bytes = f.read()
            image2_ext = os.path.splitext(image2_path)[1].lower()

            if image2_ext == '.webp':
                img = Image.open(io.BytesIO(img_bytes)).convert('RGB')
                img_bytes = io.BytesIO()
                img.save(img_bytes, format='JPEG')
                img_bytes = img_bytes.getvalue()

            slide.shapes.placeholders[first_picture_index+1].insert_picture(io.BytesIO(img_bytes))

    # Save populated PowerPoint file
    prs.save(os.path.join(app.config['UPLOAD_FOLDER'], 'mypopulated.pptx'))
    
    # Delete temporary data and picture files
    os.remove(data_path)
    for picture_path in picture_paths:
        os.remove(picture_path)
    
    # Return populated PowerPoint file for download
    return send_file(os.path.join(app.config['UPLOAD_FOLDER'], 'mypopulated.pptx'), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
