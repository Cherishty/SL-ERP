from distutils.core import setup
import py2exe

setup(
        options={'py2exe': {'bundle_files': 1,"compressed" : 1, "optimize" : 2, }},
        zipfile=None,
        console=[{"script":"ShuangLongERP.py",'icon_resources':[(1,"sl.ico")]}],

)
