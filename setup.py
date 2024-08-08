from setuptools import find_packages, setup

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

setup(
    name="excel-helper",
    version="0.1.0",
    author="Rahul Gurujala",
    author_email="isaacnewtonrahul@gmail.com",
    description="A simple library to simplify Excel manipulation using openpyxl",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/rahulgurujala/excel-helper",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 3 - Alpha",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
    ],
    python_requires=">=3.6",
    install_requires=[
        "openpyxl>=3.0.0",
        "pandas>=2.2.2",
        "Jinja2>=3.1.4",
    ],
)
