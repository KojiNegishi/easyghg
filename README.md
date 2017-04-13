# xlwings version warning
The easyghg code has been updated to work with xlwings versions 9 and up. Check your xlwings version if you encounter any errors. Log an issue if something still doesn't appear to work.

# easyghg
GHG forcing calculations in Excel using the ghgforcing Python package and xlwings. Still in very early stages, this tool will help users calculate radiative forcing from CO2 and CH4. It can account for CH4 oxidation to CO2, and can output Monte Carlo results.

## How to use
### Make sure you have Python
Right now you'll need to have Python and a few packages installed. I recommend using the [Anaconda distribution](https://www.continuum.io/downloads) if you don't already have Python. The necessary packages include:
- xlwings
- numpy
- scipy
- matplotlib
- seaborn
- ghgforcing (install by typing 'pip install ghgforcing' into a command prompt)

Everything except ghgforcing should come with the Ananconda installation. 

### Using the files
Download the easyghg.xlsm and easyghg.py files and place them in a folder together. The interface is very basic right now - let me know if you have any suggestions/requests!

On the *easyghg* sheet, enter annual CO2 and CH4 emissions into rows B & C. You don't need to enter zeros - the program will read all cells that have a value in the *Year* column.

There are a few options in row N related to CH4 (fossil CH4 oxidation to CO2, and climate-carbon feedbacks), and the number of runs for Monte Carlo calculations. Select if you are modeling a pulse emission (all emissions in the first year) for better accurary. Note that the mean value results from a Monte Carlo simulation may differ slightly from the expected deterministic value.

