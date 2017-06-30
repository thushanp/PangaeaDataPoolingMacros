# VBA-Pooling-Macros

A series of macros I built to extract, clean and visualise oil well data, pooling orders and other publicly available regulatory information from www.pangaeadata.com to assist my team at my freshman summer internship at Bowery Investment Management, LLC., a distressed debt hedge fund in New York, as they invested into physical energy assets in Oklahoma. 

Specifics:

1. Area of Interest Map Macro
	Will automatically generate a highlighted map of sections in the Kingfisher, Blaine, Canadian and Grady counties that you are interested in. This can be combined with the Datacleaner to map out areas of high regulatory activity and interest in these counties - with either large numbers of pooling orders, spud dates, active poolings or wells currently producing - which tend to be prime investment targets.
2. Datacleaner Macro
	Will clean data copied and pasted from Pangaea into a new spreadsheet organised into columns with all pertinent information listed out side-by-side in columns.
3. Colouring Macro
	Will automatically count the number of cells filled in of a specific colour - useful for calculating the total number of sections the Area of Interest Map Macro has filled in for you or to calculate how many sections are in another given area of interest.
4. CityStateZip Macro
	Simple Macro that converts strings in a single cell of form "City, State, Zip" such as "Oklahoma City, OK, 73008" into three separate adjacent cells of form "Oklahoma City", "OK", "73008"
5. Trimmer Macro
	Simple Macro that will trim the final character for a string in a cell. Useful for Area of Interest Map Macro - if you have a long list of "15"/"16N"/"8W" type data, then the trimmer can remove the "N" and "W" quickly and easily before the Area of Interest macro goes to work

Pangaea requires a login, but with a few small edits, these macros can also be utilised with the Oklahoma Corporation Commission website http://www.occeweb.com/ which is freely accessible.