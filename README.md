# TOPSIS_Solver
This programme is an implementation of TOPSIS multi-criteria decision analysis method. It returns the ranking of alternatives for given criterias and decision matrix.

## Installation
For propper operation of this programme you must have Python >=3.0 installed on your computer. For installation instructions for your OS you should go to [official Python webpage](https://www.python.org/).

Before first use of the programme you have to run an initialisation script called `init.sh` in the main directory of the programme. It installs all necessary libraries on your computer. To execute the script you should type in your terminal:
* if you're on Linux/Mac: `./init.sh`
* if you're on Windows: `init.sh`


## Usage
To use the programme you run the `topsis_solver.sh` script with propper arguments:

`./topsis_solver.sh input_file_path [output_directory]`

`input_file_path` is a path to the input file in csv format (`;` as a field separator, and `,` used for decimal numbers) containing the decision matrix as well as weights for all the criteria and information whether each criterion has a positive or negative impact.

`output_directory` (optional argument) is a path to the directory where you want the report file to be generated. If not specified the directory of the input file is used. The report contains all the alternatives and respective CI scores (similarity to the best solution), ordered by the CI score (descending). The report is generated in `.xlsx` format.


### Input file preparation
The input file should be prepared in MS Excel programme following this example (colours are only for demonstration purposes)

![image](https://user-images.githubusercontent.com/23523276/226003270-ad731927-1cb3-4b85-ab62-541862012aa5.png)

* First row (yellow colour) contains weights for each criterion (weights should add up to 1)
* Second row (blue colour) contains information whether a criterion has a positive (`+`) or negative (`-`) impact
* Fileds in green contain names for the criteria and alternatives
* Fields in white contain the decision matrix data

Such table should be saved as a `.csv` file: File &#8594; Save as &#8594; Choose a ".csv" file format (and make sure you have `;` as a field separator) &#8594; Save


