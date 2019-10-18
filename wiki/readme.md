# SurveyMerge
SurveyMerge is an Excel tool which facilitates data analysis of survey data files created by the [PIEL Survey app](http://pielsurvey.org) . It merges multiple survey runs and multiple survey files. SurveyMerge runs within Excel on Windows. At this stage it does not support Excel on Mac. It is written in Visual Basic for Applications (VBA).

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
