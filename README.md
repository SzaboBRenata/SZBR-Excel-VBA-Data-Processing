SZBR – Excel VBA Data Processing

Project description

	The project demonstrates automated Excel data processing using VBA macros. It contains two main functions: merging multiple Excel files into one worksheet and splitting a source workbook into separate files by county.

1. MergeFiles – Excel files merging

	Purpose: Combine five Excel files from the ADAT folder into one worksheet in the macro-enabled workbook.

	Steps

		1.Place the five input Excel files in the ADAT folder.

		2.Open SZBR_Excel_VBA_Data_Processing.xlsm.

		3.Run the MergeFiles macro.

		4.The macro merges the contents of the five files into the first worksheet (Munka1) of the macro-enabled workbook.

2. SplitFiles – Export by county

	Purpose: Split the source Excel file into separate workbooks based on county names.

	Steps

		1.Place forrasadatok.xlsx in the project root folder, next to the macro-enabled workbook.

		2.Make sure the source worksheet is named megye.

		3.Open SZBR_Excel_VBA_Data_Processing.xlsm.

		4.Run the SplitFiles macro.

		5.The macro creates a separate Excel file for each county.

Project structure

	SZBR_Excel_VBA_Data_Processing.xlsm – Macro-enabled Excel workbook containing both VBA macros.

	ADAT/ – Five input Excel files for the merging process.

	forrasadatok.xlsx – Input file for the county-based export.

	VBA/ – Exported VBA source modules (.bas files).

Technologies

	Microsoft Excel

	VBA (Visual Basic for Applications)

Notes

	All included input data is fictional and provided for demonstration purposes.
