# XLPlan
Code to scrape specific XL-files for data and collecting it in a specific XL-file using OpenPyXL

This project was made in correlation with a "PLAN" course (type of course for Danish scouts)
The course had the participants take a test for which of the rolls in Belbins Groupe Theory appeals to them the most. This test is seen in the 'Template.xlsx' file. Then, it scrapes the files for relevant data: Name and the scores of the different roles. Then, the program inserts the values on the 'overview' sheet, and dedicates a template sheet to the participant. 

It was requested that the data should be collected in a XL-file and that a Spider chart was created for each partisipant. However this posed a problem since OpenPyXL is not capaple to create or even interact with graphs in XL-files. This would not have been a problem for how graphs i XL references it's data cells. They reference the cells by an exact sheet (ex: "='sheet 1'!A1"). If a sheet's name is changed in XL, the graphs update automaticly, but using the program, does not update the reference when the sheet name is changed. This was worked around by 'hardcoding' each graph to be using the corresponding cells in the overview sheet.

The UI was created using TKinter. Most of the program's UI is programmed with Danish text for a Danish user.
