<!-- HEADER -->
<div align="center">
  
  # CertGen 📝

  A simple certificate generator, written in Python.

  <a href="#demonstration">View Demo</a>
  ·
  <a href="https://github.com/ihsingh2/certgen/issues">Report Bug</a>
  ·
  <a href="https://github.com/ihsingh2/certgen/issues">Request Feature</a>
</div>



<!-- TABLE OF CONTENTS -->
<details>
  <summary>Table of Contents</summary>
  <ol>
    <li><a href="#demonstration">Demonstration</a></li>
    <li><a href="#getting-started">Getting Started</a></li>
    <li><a href="#data-input">Data Input</a></li>
    <li><a href="#template">Template</a></li>
    <li><a href="#customisation">Customisation</a></li>
    <li><a href="#roadmap">Roadmap</a></li>
  </ol>
</details>

<!-- DEMONSTRATION -->
## Demonstration

| Console                                                                                 | Data                                                                                 | Output                                                                                 |
| --------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------ | -------------------------------------------------------------------------------------- |
| <img src="https://github.com/ihsingh2/certgen/blob/main/demo/console.png" width="650"/> | <img src="https://github.com/ihsingh2/certgen/blob/main/demo/data.png" width="650"/> | <img src="https://github.com/ihsingh2/certgen/blob/main/demo/output.png" width="650"/> |


<!-- GETTING STARTED -->
## Getting Started

1. Clone this repo ```git clone https://github.com/ihsingh2/certgen.git```
2. Install all dependencies from ```requirements.txt``` by using ```pip install -r requirements.txt```
3. Run main.py

<!-- DATA -->
## Data Input

CertGen accepts input for name field by the following methods:
1. CSV File
2. Text File
3. Spreadsheet
4. Console Based Input

Input in each case is stored in the form of a list, that is iterated during certificate generation.

<!-- TEMPLATE -->
## Template

CertGen supports templates in the following formats:
1. PNG
2. JPG
3. PPT
4. PDF

Template is converted to a [PIL](https://pillow.readthedocs.io/en/stable/) compatible format, in case of PPT and PDF (first slide/page only).

<!-- CUSTOMISATION -->
## Customisation

Lastly, the user is expected to enter the following details:
1. Path of font file to use
2. Color of font to use
3. Coordinates of name field
4. Text alignment of name field
5. Path of output folder

<!-- ROADMAP -->
## Roadmap

1. Prompt error handling, such as in case of wrong path.
2. Option to bulk send certificates to corresponding e-mail addresses.
3. Progress bar to display the number of certificates generated at a given time.
4. Explore similar project domains, such as marksheet generator.
