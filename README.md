# HPC Volume Calculator

A desktop application that runs equations on cell count and hippocampal areas and outputs hippocampal volume and cell density data. Written in Python 2.7, GUI built with wxPython. 

To run locally, first make sure you have the package manager [pip](https://pip.pypa.io/en/stable/installing), then make sure you have the following dependencies: 

+ [wxPython](http://www.wxpython.org/download.php)
+ [xlrt & xlwt](http://pythonhosted.org/xlutils/installation.html)
+ [ElementTree](http://effbot.org/downloads#elementtree)

Run `python main.py` to get started. 

## Purpose

Assumptions: [future versions of this app may be more flexible]
- tissue slices were collected in a series of 10 and are 40µm thick
- both sides of the brain are being analyzed

This program calculates the volume of specific hippocampal regions (based on areas from ImageJ) and cell density within the analyzed region. Raw data are based on one tenth of the hippocampus and expolations to the whole structure are made using Cavalieri's principle. 

The calculated data are entered into a new spreadsheet. Format is meant to make it easy to copy data directly into Statistica. 


