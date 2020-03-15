import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

with open("requirements.txt", "r") as req:
    reqs = req.read().splitlines()

setuptools.setup(
    name="automate_excel",
    version="0.0.1",
    author="Chris Charlton",
    author_email="chrispcharlton@gmail.com",
    description="A library for automating existing spreadsheets.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/chrispcharlton/automate_excel",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "Operating System :: Microsoft Windows",
    ],
    install_requires=reqs,
    python_requires='>=3.7',
)