#!/bin/bash

# Clear the output_tickets folder
rm -f "output_tickets/"*

python3 pdf_split.py
python3 sender.py