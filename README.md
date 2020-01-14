# QR-Gen
<p align="center">![QRGEN Interface](https://user-images.githubusercontent.com/31670621/72321516-253da980-36b5-11ea-912f-f60352544773.PNG)</p>

Quarterly Report Generator developed internally during my internship in Saudi Aramco (Which has an extremely strict IT limitation, many nights were spent installing libraries offline in Anaconda by copy and pasting). QR Gen saved our unit a week time of editing and staying overtime. Can be customized to generate your own report in seconds if it follows a consistent structure!

## Getting Started

These instructions will get you a copy of the project up and running on your local machine for development and testing purposes.

### Prerequisites

In order to run QRGEN, you need to install the dependencies first using pip. (Python >= 3.6)
```
pip install req.txt
```

### Requirements

QRGEN requires an Excel file in the same folder that contains (in any order) 
```
ID (Submission) --- Submission Type --- Departments --- Department --- Body --- Title --- Inventor
```
Note: Inventor can be many people
and an attachment folder directory that contain folders with each folder representing a submission ID
* Employee's photo image named after the employee's name (Or close enough if Fuzzy Matching is enabled)
* Invention idea images named in numeric sequence starting from 1 (1, 2, 3, ...)
structured in the following way
```
    ├── attachments                 
    │   ├── 1          # Employee IDs
            └── Asaad AlGhamdi.jpg    # Attachments
            └── Aisha AlMulla.jpg    # Attachments
            └── 1.jpg
            └── 2.png
            └── 3.jpg
    │   ├── 2
            └── Amjad AlKhathlan.jpg 
            └── 1.jpg 
            └── 2.jpg 
    │   └── ...                
```

Once you have the following requirements met, you 


## Usage

In order to start QRGEN, run in the cmd:
```
python UI.py
```

You have to select an Excel DB file, the attachment directory and check the desired optional features before pressing Generate. The output are Word Document files that are divided into two types (Idea or Success Story).
## Built With

* [Python-docx](https://python-docx.readthedocs.io/en/latest/) - The Word Document generator library
* [FuzzyWuzzy](https://pypi.org/project/fuzzywuzzy/) - Fuzzy Matching attachment names with inventor names
* [Pandas](https://pypi.org/project/pandas/) - Handling Excel Files (It's like bringing a machine gun to a knife fight, I know)
* [dlib](https://pypi.org/project/dlib/) - Cropping employee faces out of images
* [PyQt5](https://pypi.org/project/PyQt5/) - QRGen User Interface

## Acknowledgments

I'd like to thank Amjad & Aisha for their support in delivering this project. This project along with two other internal projects have caused cost savings of $500,000 annually by streamlining Saudi Aramco's innovation unit business processes.

