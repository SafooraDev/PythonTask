import csv
import os
from pptx import Presentation
from pptx.util import Inches

def load_csv(csv_file):
    data = {}
    with open(csv_file, mode='r') as file:
        reader = csv.DictReader(file) #this will create a csv reader object that will read the csv file as a dictionary 
        for row in reader:
            #print("I am inside the for loop")
            data[row['Key']] = row['Data'] #add the key and data from each row inot the data dictionary 
    print("CSV Data:", data)  # Debugging statement
    return data

def update_pptx(pptx_file, data, output_file):
    prs = Presentation(pptx_file)
    for slide in prs.slides:
        for shape in slide.shapes: #iterate through each slide in my pptx
            if shape.has_text_frame: #iterate through eavh shape in my pptx
                for paragraph in shape.text_frame.paragraphs: # if my shape contains text
                    for run in paragraph.runs: #iterating through each paragraph in the text frame
                        key = run.text.strip() #iterate through each run of the text in the paragraph
                        if key in data: #checks if key is in the data dictionary 
                            run.text = data[key] # now update the text in run with the corresponding value in the dictionary 
                            print(f"Updated text for {key}: {data[key]}")  # Debugging statement
            elif shape.shape_type == 13:  # if the shape is a Picture
                #print("i am insie the picture else if")
                key = shape.name.strip() #get the name of the shape
                #print(key)
                #print(data)
                
                if key in data: #if the key is in the dictionary
                    #print("I am inside the if statement for shape.name")
                    image_path = data[key] #check the image form the dictionary
                    if os.path.exists(image_path):
                        shape._element.getparent().remove(shape._element)  # Remove old image
                        slide.shapes.add_picture(image_path, shape.left, shape.top, shape.width, shape.height)
                        print(f"Updated image for {key}: {image_path}")  # Debugging statement
                    else:
                        print(f"Image path does not exist: {image_path}")  # Debugging statement
    prs.save(output_file)

# Full paths to the files on your Desktop using raw strings
csv_file = r'/home/dev/Desktop/data.csv'
pptx_file = r'/home/dev/Desktop/source_presentation.pptx'
output_file = r'/home/dev/Desktop/updated_presentation.pptx'


# Ensure the CSV and PPTX files exist
if not os.path.exists(csv_file):
    raise FileNotFoundError(f"The file {csv_file} does not exist.")
if not os.path.exists(pptx_file):
    raise FileNotFoundError(f"The file {pptx_file} does not exist.")

# If the output file already exists, remove it to avoid permission issues 
if os.path.exists(output_file):
    os.remove(output_file)

# Load data from CSV and update PPTX
csv_data = load_csv(csv_file)
update_pptx(pptx_file, csv_data, output_file)