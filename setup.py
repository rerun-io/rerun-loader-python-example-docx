from setuptools import setup

# Read requirements from the requirements.txt file
with open("requirements.txt") as f:
    requirements = f.read().splitlines()

setup(
    name="rerun-loader-python-example-docx",
    version="0.1.0",
    py_modules=["rerun_loader_docx"],
    python_requires=">=3.8",
    install_requires=requirements,
    entry_points={
        "console_scripts": [
            "rerun-loader-docx=rerun_loader_docx:main",
        ],
    },
)
