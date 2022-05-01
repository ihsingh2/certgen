from pandas import read_csv, read_excel
from PIL import Image, ImageDraw, ImageFont
from pdf2image import convert_from_path
from tempfile import gettempdir
import win32com.client

print("──────────────────────────────────────── WELCOME TO CERTGEN! ───────────────────────────────────────")
print("Author: Himanshu Singh")
print("Last Modified: 30-04-2022")

temp_dir = gettempdir()
i, j = 0, 0

def names():
    '''Takes input for the certificate's name field and returns it in the form of a list.'''

    #counter for displaying the help info only once
    global i
    if (i == 0):
        print("────────────────────────────────────── STAGE 1/3 - DATA INPUT ──────────────────────────────────────")
        print("1: CSV File")
        print("2: Text File")
        print("3: Spreadsheet")
        print("4: Manually")
    choice = int(input("Select a method of data input for name field: "))

    #reads the csv file as a dataframe and stores values of the user-specified row in a list
    if (choice == 1):
        path = input("Enter the path of spreadsheet to import data from: ")
        data = read_csv(path)
        col = input("Enter the column header to extract data from: ")
        lst = data[col].tolist()

    #reads the txt file and stores content of each line in a list
    elif (choice == 2):
        path = input("Enter the path of text file to import data from: ")
        lst = []
        with open(path, "r") as file:
            for line in file:
                lst.append(line.split('\n')[0])

    #converts the spreadsheet to a csv file and stores values of the user-specified row of csv read dataframe in a list
    elif (choice == 3):
        path = input("Enter the path of spreadsheet to import data from: ")
        file = read_excel(path)
        path = temp_dir+"\data.csv"
        file.to_csv(path, index = None, header=True)
        data = read_csv(path)
        col = input("Enter the column header to extract data from: ")
        lst = data[col].tolist()
    
    #takes 'num' inputs from the user, and stores them in a list
    elif (choice == 4):
        num = int(input("Enter the number of certificates to generate: "))
        lst = []
        for i in range (num):
            name = input("Enter name "+str(i+1)+": ")
            lst.append(name)

    #repeats the input statement till the user has entered a valid input
    else:
        print("Please enter an integer from 1 to 4.")
        i += 1
        return names()

    #returns the list of names
    return lst

def template():
    '''Takes input for the certificate's template and returns it in a PIL compatible format.'''

    #counter for displaying the help info only once
    global j
    if (j == 0):
        print("──────────────────────────────────── STAGE 2/3 - TEMPLATE INPUT ────────────────────────────────────")
        print("1: PNG")
        print("2: JPG")
        print("3: PPT")
        print("4: PDF")
    choice = int(input("Select the format of certificate template to use: "))
    
    #stores the path of png file as it is
    if (choice == 1):
        tmp = input("Enter the path of template: ")
    
    #stores the path of jpg file as it is
    elif (choice == 2):
        tmp = input("Enter the path of template: ")

    #converts the first slide of ppt to png and stores it in the user's temporary directory
    elif (choice == 3):
        tmp = input("Enter the path of template: ")
        Application = win32com.client.Dispatch("PowerPoint.Application")
        Presentation = Application.Presentations.Open(tmp)
        tmp = temp_dir+"\template.png"
        Presentation.Slides[0].Export(tmp, "PNG")
        Application.Quit()
        Presentation = None
        Application = None
    
    #converts the first page of pdf to png and stores it in the user's temporary directory
    elif (choice == 4):
        tmp = input("Enter the path of template: ")
        images = convert_from_path(tmp)
        tmp = temp_dir+"\template.png"
        images[0].save(tmp, "PNG")
        pass

    #repeats the input statement till the user has entered a valid input
    else:
        print("Please enter an integer from 1 to 4.")
        j += 1
        return template()

    #returns the path of the template image file
    return tmp

def custom():
    '''Takes input for font file, coordinates of name field and returns the same in the form of a list.'''

    print("────────────────────────────────── STAGE 3/3 - TEXT CUSTOMISATION ──────────────────────────────────")
    font = input("Enter the path of font file to use: ")
    color = input("Enter the color to use (name, #rgb, rgb(r,g,b), hsv(h,s%,v%)): ")
    x1 = int(input("Enter the x coordinate of the left edge of name field: "))
    x2 = int(input("Enter the x coordinate of the right edge of name field: "))
    y1 = int(input("Enter the y coordinate of the upper edge of name field: "))
    y2 = int(input("Enter the y coordinate of the lower edge of name field: "))
    align = input("Enter the text alignment of name field (left, center, right): ")
    print('─' * 100)
    lst=[font, color, x1, x2, y1, y2, align]
    return lst

def generate(names: list, template: str, font_path: str, color: str, x1: int, x2: int, y1: int, y2: int, align: str):
    '''Generates certificates for a list of names, provided an image template, font file, color, coordinates of name field and text alignment.'''

    path = input("Enter the path of output folder: ")
    
    for name in names:
        
        #opens the template image and calculates the dimensions of text for default font size
        img = Image.open(template)
        field_width = x2 - x1
        font_size = y2 - y1
        draw = ImageDraw.Draw(img)
        font = ImageFont.truetype(font_path, font_size)
        text_width, text_height = draw.textsize(name, font = font)

        #reduces font size for names longer than field_width
        k = 0
        while (text_width > field_width):
            font_size = font_size - 5
            old_text_height = text_height
            font = ImageFont.truetype(font_path, font_size)
            text_width, text_height = draw.textsize(name, font = font)
            k = k + old_text_height - text_height

        #correction factor for upper y coordinate for long names
        m = k/6
        y3 = y1 + k - m

        #final processing of image
        if (align == "left" or align == "Left" or align == "LEFT"):
            draw.text((x1, y3), name, fill=color, font = font)
        elif (align == "right" or align == "Right" or align == "RIGHT"):
            draw.text((x2 - text_width, y3), name, fill=color, font = font)
        else:
            draw.text((x1 + (field_width - text_width)/2, y3), name, fill=color, font = font)

        #saves the processed image in the output folder
        img.save(r"{}\{}.png".format(path,name))

#list of names
names = names()

#string path of editable template file
template = template()

#list of user preferences for text field
custom = custom()

#passing the names, template path and user preferences for text field to the generate function
generate(names, template, custom[0], custom[1], custom[2], custom[3], custom[4], custom[5], custom[6])

print("Certificates generated successfully!")

exit = input("\nPress any key to exit...")