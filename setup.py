from setuptools import find_packages, setup

with open("README.md", "r", encoding="utf-8") as fh:
    long_description = fh.read()

with open("requirements.txt", "r", encoding="utf-8") as fh:
    requirements = fh.read().splitlines()

setup(
    name="gridient",
    version="0.1.0",
    author="Your Name",  # TODO: Replace with your name
    author_email="your.email@example.com",  # TODO: Replace with your email
    description="A Python library for writing calculations to Excel while preserving formulas.",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/gridient",  # TODO: Replace with your repo URL
    packages=find_packages(),
    install_requires=requirements,
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",  # Choose your license
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.7",  # Specify your minimum Python version
)
