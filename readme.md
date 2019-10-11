# SurveyMerge
SurveyMerge is an Excel tool which facilitates data analysis of survey data files created by the [PIEL Survey app](http://pielsurvey.org) . It merges multiple survey runs and multiple survey files. SurveyMerge runs within Excel on Windows. At this stage it does not support Excel on Mac and must be regarded as a Beta version. It is written in Visual Basic for Applications (VBA).

The PIEL Survey app is a research tool which conducts surveys using Ecological Momentary Assessment Methodology (EMA), otherwise known as Experience Sampling Methodology (ESM). It is available as an iOS and Android app.

- [PIEL Survey on the App Store](https://itunes.apple.com/au/app/piel-survey/id1257313392?mt=8)
- [PIEL Survey on Google Play](https://play.google.com/store/apps/details?id=au.com.bluejay.pielsurvey)

The PIEL Survey app is free and used widely in research and therapy contexts. Please consider ["forking"](https://help.github.com/en/articles/fork-a-repo) the project and helping to improve it for all users.

## Installation
You can download an Excel file with the VBA code of the latest release pre-installed from [the PIEL Survey website](http://pielsurvey.org/download/surveymerge). However, we expect that many organizations will want to review the code and install it themselves.

#### End User
1. Download a zip file of the files in  the repository.
2. Create a new Excel spreadsheet and save it as an "Excel Macro-Enabled Wookbook (*.xlsm).
3. Open the Visual Basic Editor: "Alt-F11".
3. Import all the files in the `/src` folder into the VBAProject.
4. Open the `Main` module.
5. Place the curser in the the subroutine `install()` run it: "Alt-F5". This will create the required worksheets.
   - The dashboard worksheet which contains the instructions and merge button for the end user.
   - The "Answer" worksheet which has the main data for the survey.
   - The "Answer Time" which has the times each question was answered.
6. Distribute the spreadsheet to the end user.

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
