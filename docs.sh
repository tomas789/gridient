#!/bin/bash

# Navigate to the docs directory
cd docs

# Install documentation dependencies
python -m pip install -r requirements.txt

# Build the HTML documentation
make html

# Navigate back to the root directory
cd ..

# Print the location of the built documentation
echo "Documentation built successfully. Open docs/_build/html/index.html in your browser." 