from setuptools import setup, find_packages
import sys, os

version = '1.0.0'

setup(name='Lazy-eCert',
      version=version,
      description="Stupid Cli app for bulk e-certificate generator and deployer via email.",
      long_description="""\
""",
      classifiers=[], # Get strings from http://pypi.python.org/pypi?%3Aaction=list_classifiers
      keywords='Certificate Generator',
      author='2k16daniel',
      author_email='2k16daniel@gmail.com',
      url='https://github.com/2k16daniel/Lazy-eCertGen',
      license='MIT',
      packages=find_packages(exclude=['ez_setup', 'examples', 'tests']),
      include_package_data=True,
      zip_safe=True,
      install_requires=[
           'CLick',
           'numpy',
           'openpyxl',
           'pandas',
           'PasteScript',
           'pillow',
           
      ],
        entry_points={
            'console_scripts': [
                'eCert = lazyecert.eCert:cli',
            ],
        },
      )
