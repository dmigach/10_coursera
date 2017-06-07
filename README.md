# Coursera dump

This script allows you to dump info about courses from Coursera to .xlsx spreadsheet.

## Prerequisites

The script is written in `Python 3`, so you'll need it's interpretator to run it.

## Install

To install all the necessary libraries to run the script just open your terminal, go to downloaded project directory and type:

    pip install -r requirements.txt

## Usage

To run the script type following in terminal:
    
    python coursera.py
    

By default you'll get a spreadsheet with info about 20 courses, but you can customize script run by adding arguments to call in form of

    python coursera.py courses_amount output_filename name_coilumn_width language_coilumn_width
    
By default output file name is `courses.xlsx` and cutomizable columns width is set to `40`.

For example, if you want to get 5 courses, widen both columns twice and output info to my_courses.xlsx you should run

    python coursera.py 5 my_courses 80 80
    
To get arguments help type

	python coursera.py -h 

## Support

In case of any difficulties or questions please contact <dmitrygach@gmail.com>.