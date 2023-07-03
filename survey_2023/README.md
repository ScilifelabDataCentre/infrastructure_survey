## Survey 2023

This folder contains the scripts used for Scilifelab infrastructure survey 2023

#### single_survey_page.py

This script takes the survey output (an Excel file provided by Scilifelab Operations Office), and creates individual pdf file for each response. The output will saved in a folder called `Pdfs` (which will contain sub folders in the name of Scilifelab platform that the user selected in the survey)

**Usage:**

```
python single_survey_page.py
```

#### Make_plots.py

This script takes the survey output (an Excel file provided by Scilifelab Operations Office), and creates summary plots and statistics of the responses. The plots will be saved in a folder called `Plots`. In total, there are 4 types of barplot and 7 individual plots. Three types of plot are created for both types of survey (A & B):

- Affiliation of the respondant (plots produced are affiliation_A.svg and affiliation_B.svg).

- Which SciLifeLab Platform would the suggestion fit into? (plots produced are platform_fit_A.svg and platform_fit_B.svg).

- Which SciLifeLab capability would the suggestion fit into? (plots produced are capability_fit_A.svg and capability_fit_B.svg).

One type of plot is only created for survey type B:

- The estimated number of unique annual visitors if the facility was integrated into SciLifeLab's national infrastructure (plot produced is potential_users_B.svg). The colour of the bars on the graph corresponds to the colour selected for the headers in pdf documents created for that survey type (either A or B).

**Usage:**

```
python Make_plots.py
```

#### Make_graphs_pdfs.py

This script imports the plots and summary statistics generated in the `Make_plots.py` script and integrates them into a pdf file. The text in the file is coloured according to the survey type, and the colour is the same as that used for the bars in the plots. The phrasing of the headers differs according to survey type. The output is saved in a folder called `pdfs_plots`.

**Usage:**

```
python Make_graphs_pdfs.py
```
