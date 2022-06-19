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
    """Stupid Cli app for bulk e-certificate generator and deployer via email."""
    # check first if config file exists
    if config_file is None:
        print('Config file is missing, skipping...')
    else:
        print(config_file)
    # load names from excel
    Name = excel_parser(excel_file)

    certGenerator(Name,font,template,coordinate_x,coordinate_y,filename,out_dir)


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
        for i in range(1,6):
            certificate = Image.open(template)
            draw = ImageDraw.Draw(certificate)
            myfont = ImageFont.truetype(font,size=148)
            color = 'rgb(45, 52, 54)'
            name , email = entries[i]
            print('Printing certificate for :' + name)
            draw.text((_x, _y), name, fill=color, font=myfont)
            os.makedirs(output_path + "/" + email,exist_ok=True)
            certificate.save(output_path + "/" + email +"/" + filename + '_'+ name + '.png')

            del draw
    except:
        traceback.print_exc()