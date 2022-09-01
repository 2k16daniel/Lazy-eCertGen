import click, os, openpyxl, traceback
from click.types import STRING
from PIL import Image, ImageDraw, ImageFont

# Optional lists


@click.command()
@click.option('--config-file', type=click.File(), help='Path of your custom config file')
@click.option('--filename',  help='Path of your custom config file')
@click.option('--template', type=click.Path(), help='Path of your eCertificate template ( Must be in image format)')
@click.option('--font', type=click.Path(), help='Path of your eCertificate template ( Must be in image format)')
@click.option('--out-dir', default="output", type=click.Path(), help='Your desired output path')
@click.option('--excel-file', type=click.Path(), help='Path of your excel file')
@click.option('--coordinate-x', type=int, help='Coordinates on the certificate where will be printing the name ')
@click.option('--coordinate-y', type=int, help='Coordinates on the certificate where will be printing the name ')

def cli(out_dir, excel_file, config_file, template, coordinate_x, coordinate_y, font, filename):
    """Stupid python Cli app for bulk generating of e-Certificates"""
    # check first if config file exists
    if config_file is None:
        print('Config file is missing, skipping...')
    else:
        print(config_file)
    # load names from excel
    Name = excel_parser(excel_file)
    certGenerator(Name,font,template,coordinate_x,coordinate_y,filename,out_dir)
    exit()

def excel_parser(excel_file):
    print('fetching names from your excel file...')
    obj = openpyxl.load_workbook(excel_file)
    sheet = obj.active
    Name = []
    # calculate num of rows that has value
    nrow_max = len([row for row in sheet if not all(
        [cell.value == None for cell in row])])
    print('Max entries : ',nrow_max)
    #store all email and name into array (tupple)
    for i in range(1, nrow_max+1):
        Name.append([''.join(
            [
                sheet.cell(row=i, column=1).value
            ]
        ),
            sheet.cell(row=i, column=2).value])
    return Name

def certGenerator(entries, font, template, _x, _y, filename, output_path):
    print('Creating folder for generated output...')
    print(template)
    if os.path.exists(output_path):
        print("Folder already exists, merging file instead")
    else:
        os.makedirs(output_path,exist_ok=True)
    try:
        for i in range(0,len(entries)):
            certificate = Image.open(template)
            draw = ImageDraw.Draw(certificate)
            name , email = entries[i]

            myfont = ImageFont.truetype(font,size=textAdjust(name))
            color = 'rgb(245, 167, 19)'
            print('Generating certificate for : ' + name)
            new_X , new_Y = coordinateAdjust(name,_x,_y)
            draw.text((new_X,new_Y), name, fill=color, font=myfont)
            os.makedirs(output_path + os.sep + email,exist_ok=True)

  
            certificate.save(os.path.join(output_path, email+(os.sep), filename + '_' + goAwayWhitespaceDamnit(name) + '.png'), quality= 85, optimize=True)

            del draw
    except:
        traceback.print_exc()

def textAdjust(name):
    max_length = 30
    max_fontsize = 188

    curr_length = len(name)
    return max_fontsize - ((curr_length - max_length) * 2)

def coordinateAdjust(name,x,y):
    max_length = 21
    curr_length = len(name)


    new_x =  x - ((curr_length - max_length) * 20)
    new_y =  y + ((curr_length - max_length) * 2)

    return new_x,new_y

def goAwayWhitespaceDamnit(name):
    return name.replace(' ', '-')