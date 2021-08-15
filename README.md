# Arkema OCR Implemenation
Team members: Prats Mathur

## Table of Contents
* [Project Status](#project-status)
* [Project Purpose](#project-purpose)
* [Project Constraints](#project-constraints)
* [Technologies Used](#technologies-Used)
* [Setup](#setup)

## Project Status
Dead.
After management review, this project has been pivoted to the Power Automate platform to maintain a homogenous program portfolio for the company.

### To Do
* [X] Create a script that can take an arbitray number of PDFs and textsearch them
* [X] Implement an indexing and file storage system to make PDFs easily findable
* [ ] Generalize implementation to search for any set of key values
* [ ] Implement auto-timed purge functionality

## Project Purpose
This project was created to remedy a complete wasteage of resources I'd noticed as an intern- Filing Paperwork. Specifically, by the ISO9000 standard, our site was required to hold on to all work orders for at least 1 year or, if it tested a Safety Critical Safeguard (SCS), for 3 years. This led to each years pair of interns spending about a day every month sorting, filing and purging work orders from the large filing cabinet in the engineering office.

I figured if I implemented a quick python prototype that would take the scanned files (scanned as operators created them, not all at once like the old system) I could create a business case for management to fund a proper implementation. 

As I had been doing this project on an interns budget, Google's Tesseract OCR engine had all of the traits I was looking for, namely that it could be accessed via simple API calls and was free.

## Project Constraints
 
* Must process an arbitrary number of PDF scans - determine procedure type, assign ID, index as needed and store file as required
* Must be faster than manually processing the files
* Must be useable by someone with limited technological knowledge (this specific person may or may not still use IE as their main browser)
* Must keep files local to Arkema machines to avoid any information leaks.
* Must be able to run without a python interpretter (none on work machines)

## Technologies Used

### Python 3.XX (I used 3.5 for this project)
    > Would other languages be faster? Probably, but this was meant to be a quick prototype to later be transitioned into an
    > actual compiled program
#### [Pandas](https://pandas.pydata.org/)
    > Pandas dataframes were used to expedite processing of the document index file
#### [PDF2Image](https://pypi.org/project/pdf2image/)
    > This python module was used to convert scaned PDFs to jpgs to better feed into the Tesseract OCR engine
#### [Openpyxl](https://openpyxl.readthedocs.io/)
    > This python module was used as a hidden dependency to help pandas process excel files

### [Tesseract OCR](https://github.com/tesseract-ocr/)
    > Developed by Google, this is a great open-source OCR engine that is easy to use, optimizable and can even be 
    > retrained with your group-specific datasets

### [Poppler](https://poppler.freedesktop.org/)
    > A PDF rendering library, this was required for the Tesseract OCR engine.

## Setup

None required, download the "arkema_ocr.py" file alongside the poppler_home files, modify the python script with your poppler filepath and run directly with any python interpretter.
