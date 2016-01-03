import setuptools
from distutils.core import setup

setup(
    # Application name:
    name="pyxmlss",

    # Version number (initial):
    version="0.1.0",

    # Application author details:
    author="Parth Jadeja",
    author_email="parthjadeja.2001@gmail.com",

    # Packages
    packages=["pyxmlss"],

    # Include additional files into the package
    include_package_data=True,
		
		# Keywords
		keywords="xmlss reportgeneration",
    # Details
    url="http://pypi.python.org/pypi/MyApplication_v010/",

    #
    license="LICENSE.txt",
    description="Simple Report Generation tool in xmlss format.",

    long_description=open("README").read(),

    # Dependent packages (distributions)
    install_requires=[
        "dateutil",
				"lxml"
    ],
)
