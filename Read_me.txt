Update 6/7/23:
- This script is only functional for years 2016+.  Previous years have completely different formatting and 
	naming schemes, making this script unable to find what it needs.
- The program is functional, though there may still be small errors I overlooked.  Keep an eye on the box
	that shows the lines of code in case errors appear.  If an error appears, the program may stop.
- The formatting portion of the script takes a long time because it has to evaluate if a row is empty to
	decide to delete it or not.  With 23,500 rows (2017) at 19 columns per row, this part may take hours.
	The script can be closed, but it will not save the rows that are already deleted and will have to be restarted
	from the formatting step.


- The till color from the conditional formatting rules cannot be copied, so all cells will be white (minus title cells)

This works by specifying an excel file to copy everything into (template provided), and having a folder where
	all files to be copied are.  The folder must have subfolders for each month.  So folder > subfolders > files
