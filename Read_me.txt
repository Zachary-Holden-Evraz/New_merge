Update 6/7/23:
- Included a manual save button.  This saves a new file under a slightly different name (added backup to the end).
	If the script is closed before it is finished, it will be better to use the "backup" as it was more recently
	saved/updated.  
	- This does have a potential to cause a minor error in the excel file, but the error is easily cleared by
	opening the file and resaving it.  Just in case, this should only be used if you are about to leave for the
	day and cannot keep the script running.
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
