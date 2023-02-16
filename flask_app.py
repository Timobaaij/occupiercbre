from flask import Flask, render_template, request, send_file
import os
from pptx import Presentation
import pandas as pd
from PIL import Image
import io
from pptx.enum.shapes import MSO_SHAPE


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
        slide.shapes.placeholders[14].text = str(row['Adress'])
        slide.shapes.placeholders[10].text = str(row['Site size'])
        slide.shapes.placeholders[11].text = str(row['Warehouse size'])
        slide.shapes.placeholders[12].text = str(row['Office size'])
        slide.shapes.placeholders[13].text = str(row['Mezzanine size'])
        slide.shapes.placeholders[17].text = str(row['Parking'])
        slide.shapes.placeholders[18].text = str(row['Environment category'])
        slide.shapes.placeholders[19].text = str(row['Maximum building height'])
        slide.shapes.placeholders[20].text = str(row['Clear height'])
        slide.shapes.placeholders[21].text = str(row['Floor load'])
        slide.shapes.placeholders[22].text = str(row['Floor flatness'])
        slide.shapes.placeholders[23].text = str(row['Loading docks'])
        slide.shapes.placeholders[24].text = str(row['Overhead doors'])
        slide.shapes.placeholders[25].text = str(row['Sprinkler'])
        slide.shapes.placeholders[26].text = str(row['Warehouse Price'])
        slide.shapes.placeholders[27].text = str(row['Office Price'])
        slide.shapes.placeholders[28].text = str(row['Mezzanine Price'])
        slide.shapes.placeholders[29].text = str(row['Parking Cost'])
        slide.shapes.placeholders[31].text = str(row['Comments'])
        
        #Picture Google maps link
        picture_placeholder = slide.shapes.placeholders[33]
        picture_path = os.path.join(app.config['STATIC_FOLDER'], 'picture.png')
        picture = picture_placeholder.insert_picture(picture_path)
        
        # Add a transparent shape over the picture and set the hyperlink for it
        hyperlink_shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, picture.left, picture.top, picture.width, picture.height)
        hyperlink_shape.fill.background().transparency = 100000
        hyperlink_shape.click_action.hyperlink.address = str(row['Google Maps'])
        hyperlink_shape.z_order = picture.z_order + 1

        # Calculate the index of the first picture on this slide
        first_picture_index = 15

        # Construct image paths based on image filenames from Excel file
        image1_filename = str(row['Image 1'])
        image1_path = os.path.join(app.config['UPLOAD_FOLDER'], image1_filename + '.jpg')
        image2_filename = str(row['Image 2'])
        image2_path = os.path.join(app.config['UPLOAD_FOLDER'], image2_filename + '.jpg')

        # Check if image paths are valid
        if not os.path.isfile(image1_path):
            print(f'Error: Could not find {image1_path}')
        else:
            with open(image1_path, 'rb') as f:
                img_bytes = f.read()

            slide.shapes.placeholders[first_picture_index].insert_picture(io.BytesIO(img_bytes))

        if not os.path.isfile(image2_path):
            print(f'Error: Could not find {image2_path}')
        else:
            with open(image2_path, 'rb') as f:
                img_bytes = f.read()

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
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
