# Areas Calculator

*Status: user testing*

## Purpose

Assumptions: [future versions of this app may be more flexible]
- brain slices were collected in a series of 10 and are 40µm thick
- both sides of the brain are being analyzed

This program calculates the volume of specific hippocampal regions (based on areas from ImageJ) and cell density within the analyzed region. Raw data are based on one tenth of the hippocampus and expolations to the whole structure are made using Cavalieri's principle. 

The calculated data are entered into a new spreadsheet. 

Written in Python 2.7, GUI built with wxPython. 

