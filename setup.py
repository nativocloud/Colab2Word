from setuptools import setup, find_packages

setup(
    name="colab2word",
    version="0.1.0",
    packages=find_packages(),
    install_requires=[
        "python-docx>=0.8.11",
        "pandas>=1.3.0",
        "numpy>=1.21.0",
        "matplotlib>=3.4.0",
        "seaborn>=0.11.0",
        "ipython>=7.0.0"
    ],
    author="NativoCloud",
    description="A tool for generating Word documents from Google Colab notebooks",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/nativocloud/Colab2Word",
    python_requires=">=3.7",
)