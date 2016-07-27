from distutils.core import setup
import py2exe

setup(
    console=['Operautomation 2.0.py'],
    options={
            "py2exe":{
                    "packages": ["PyPDF2", "selenium", "pandas", "xlrd", "html"]
            }
            }
    )
