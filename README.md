# IMAIEF Newsletter Automation
Automation of IMAIEF analysis newsletter for Jalisco data created with R and VBA.

## Overview
In order to automate the elaboration of a newsletter that is published monthly to save time and avoid typos, a solution was developed so that the entire newsletter could be produced without having to write or make graphics, requiring no more than running a code in R, adding a macro to the generated Excel file and running it. This generates a Word document with the wording of the analysis and the graphs generated according to the latest information published by INEGI for the Monthly Indicator of Industrial Activity by State (IMAIEF).\
The newsletter was written in Spanish for the Institute of Statistical and Geographic Information of Jalisco (IIEG).


## Technologies Used
- R
- VBA Macros


## How to use the Files
1. Download the necessary R packages.
2. Change the directory in R archive to one fo your choice.
3. Run the entire R code (Ctrl A + Ctrl Enter).
4. Open the Excel file created by the code in the previously specified directory.
5. In the Excel file Go to Programmer > Code > Visual Basic.
6. In the VBAProject area right click and select "Import file" and select the Macro (.bas file).
7. Change the directory in the .bas file where the created document will be saved and the locations where the Charts and Word templates will be taken from.
8. Run the macro with Ctrl + A and it will automatically make the graphs, which will be taken to a Word document where it will write all the analysis, titles, sources and notes in the appropriate format.
