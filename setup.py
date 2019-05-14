# coding: utf-8

from setuptools import setup, find_packages
<<<<<<< HEAD
from pathlib import Path
=======
>>>>>>> 4a6c4a515a33824834296068c5176f7286f16e38

NAME = 'pandas_2_pptx'
VERSION = '0.1.0'


<<<<<<< HEAD
def set_init_dir():

    dir_list = [
        'lib',
    ]

    file_dict = {}

    symlink_dict = {
        f'lib/{NAME}': 'src',
    }

    for dir_name in dir_list:
        p = Path(dir_name)
        if p.exists():
            continue
        p.mkdir(parents=True, exist_ok=True)
        print('make dir {}'.format(dir_name))

    for file_name, file_content in file_dict.items():
        p = Path(file_name)
        if p.exists():
            continue
        print('make file {} << {}'.format(file_name, file_content))
        p.write_text(file_content)

    for source, target in symlink_dict.items():
        source_p = Path(source)
        target_p = Path(target)
        if source_p.exists() or not target_p.exists():
            continue
        for i in range(len(source_p.parts) - 1):
            target_p = Path('../').joinpath(target_p)
        print('make symbolic link {} to {}'.format(source_p, target_p))
        source_p.symlink_to(target_p)


set_init_dir()

=======
>>>>>>> 4a6c4a515a33824834296068c5176f7286f16e38
setup(
    name=NAME,
    version=VERSION,
    install_requires=["pandas", "python-pptx"],
    packages=find_packages('lib'),
    package_dir={'': 'lib'},
)
