
#### Developers
1. Either download the zip file or clone the project.
2. Follow the instructions 1-5 above.
3. A further 2 steps are needed to run the tests.
   - Import all the files in the `/testing` folder into VBAProject.
   - Create a folder `/testing/test-files` in the same folder where the workbook was saved and copy there all the contents of `/testing/test-files` in the repository.

## Working with the Code

#### Output
Currently the data output is in the form of 2 tables in 2 worksheets: "Answers" and "Answer Time". It is expected that users will modify and add to these tables.

#### Survey Data Files
SurveyMerge parses the survey data files of the PIEL Survey. These are `csv` files with a well defined format. An overview of the structure is shown [here](https://pielsurvey.org/instructions/export-results/). The tests are a more detailed source of documentation.

#### Repository
Working with the repository requires all modules to be exported to the correct folder prior to the commit. You can use the sub `exportVisualBasicCode()` in the `Utilities` module to automate this. The Excel file itself should not be included in any push request. Please note that before exporting the modules, the Excel setting "Trust access to the VBA project object model" must be checked

#### Key Classes
It is expected that users will require more worksheets. In most cases, this can be done by using the data contained in `ModelSurveyRun` objects. `ModelSurveyRun` objects contain an `Answers` collection. All objects in the `Answers` collection implement the class `ModelAnswerBase`. They  hold the data and provide access to the data in a useful format/data type.

A starting point for modifying the output is the class module: "PrinterSurveyRun".

#### Tests
The project does not use a library for unit testing but does follow a structured approach.  All test classes are named with a prefix of "Test". They conform to the interface `ITester`.

Many of the tests read data from csv files which are provided in the `/testing/test-files` folder.

To run tests currently in place, run the sub `runAllTests()` in the module `TestRunner`. Results are printed to the "Immediate" window.

## Future Enhancements
If users would like to add features, please either create a pull request or [let us know](https://pielsurvey.org/contact/). The following enhancements would be useful to add to the project.

#### SPSS Format
Add a worksheet specifically designed for SPSS import. This will likely include:
- headings that meet SPSS requirments
- format the date with seconds
- add multiple fields for checkbox questions

#### Summary Worksheet
Add a worksheet with summary data for the project.