# PubChem_Retriever.py

## Description

`PubChem_Retriever.py` is used to retrieve chemical compound data from various sources, e.g., PubChem and ClassyFire Batch by Fiehn Lab online databases. 

## What does the script do

The script takes `Excel`, `CSV`, or tab-separated `TXT` files as input and does the following:
1.	Asks the user what data to retrieve, e.g., InChIKey, CAS#, SMILES, DTXSID, Uses, and Use Classification from PubChem and chemical class data from ClassyFire Batch, through a dialog in Visual Studio Code Terminal, e.g., `Would you like to retrieve InChIKeys from PubChem? 1 = Yes, 0 = No. Type your answer and press Enter:`
2.	Asks the user to specify the input file. 
3.	Asks the user to specify input column (identifier) for data retrieval among `"Name"` and `"InChIKey"` columns.
4.	Retrieves the specified in p. 1 data, e.g., InChIKey, CAS#, SMILES, DTXSID, Uses, and Use Classification, from PubChem webpage based on the identifier specified in p. 3. If the first identifier, e.g., name, returns nothing, the second identifier, e.g., InChIKey, is used for data retrieval.
5.	Creates `"InChIKey_Consensus"` column and fills it with InChIKey values from PubChem; if an InChIKey value from PubChem is missing, fills the gaps with initial InChIKey values, e.g., from NIST23 MS library.
6.	If chosen in p. 1, retrieves chemical class data from ClassyFire Batch webpage based on `"InChIKey_Consensus"` column values.
7.	Saves and formats the output `Excel` file, including preventing CAS# values from being formatted as date in `Excel`.

## Prerequisites

Before using the script, several applications/tools have to be installed:
1.	Visual Studio Code; https://code.visualstudio.com/download.
2.	Python 3; https://www.python.org/downloads/windows/.
3.	Python Extension in Visual Studio Code > Extensions (`Ctrl + Shift + X`) > Search “python” > Press `Install`.

Then, the packages must be installed as follows:
Visual Studio Code > Terminal > New Terminal > In terminal, type `pip install package_name`, where `package_name` is a desired package name, e.g., `numpy` > Press `Enter`.

## How to use the script

To use the script, the following steps must be executed:
1.	Right mouse click anywhere in Visual Studio Code script file > Run Python > Run Python File in Terminal or press `play` button in the top-right corner.
2.	Choose the files for processing in the new pop-up window and press `Open`.

## Notes and recommendations

The input file must contain at least the following columns to be processed: 
`"Name"` and/or `"InChIKey"`

## License
[![MIT License](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/license/mit)

Intended for academic and research use.
