# Burnside Men's League Calculator

Calculates the average course index and handicap for players in a spreadsheet. Does this by taking the average of the lowest differentials by score. Requires an input XLSX with a Players sheet in a specific format.

## Requirements

* Python 3.6+

## Setup

`python -m pip install -r requirements.txt`

## Usage

`python calc.py input.xlsx`

Will produce `output.xlsx` when finished. You cannot have the output file open when running the script.