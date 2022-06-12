from typing import Type
from typing_extensions import Required
import click

#Optional lists
@click.command()
@click.option('--config-file',type=click.File(), help='Path of your custom config file')
@click.option('--out-dir', default="output" , type=click.Path(), help='Your desired output path')
@click.option('--excel-file', type=click.File(), help='Path of your excel file')


def cli(out_dir,excel_file,config_file):
    """Stupid Cli app for bulk e-certificate generator and deployer via email."""
    #check first if config file exists
    if config_file is None: 
        print('Config file is missing, skipping...')
    else: print(config_file)
    #load names from excel
    excel_parser(excel_file)

def excel_parser(excel_file):
    print('fetching names from your excel file...');
