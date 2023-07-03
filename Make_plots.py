"""This script produces the plots required for the survey data"""

import pandas as pd
import os
import plotly.graph_objects as go
import plotly.express as px
import numpy as np

# from colour_science_2023 import (
#     SCILIFE_COLOURS,
#     FACILITY_USER_AFFILIATION_COLOUR_OFFICIAL_ABB,
# )

# read in data and perform data preparation
# First portion of script (before splitting for survey type is general survey prep)

survey_data_raw = pd.read_excel(
    "Data/test_data.xlsx",
    sheet_name="Sheet 1 - 230614100957_scilifel",
    header=0,
    engine="openpyxl",
    keep_default_na=False,
)

# make affiliations types into a unified column
# (prep for affiliations work)

survey_data_raw["Affiliation"] = np.where(
    (survey_data_raw["Affiliation"] == "University"),
    survey_data_raw["University"],
    survey_data_raw["Affiliation"],
)

# Rename columns needed to work with

survey_data_raw.rename(
    columns={
        "In which of the existing SciLifeLab Platform(s) would the technology/instrument/service/technological capability fit. https://www.scilifelab.se/services/infrastructure-organization/": "Tech_fits",
        "In which of the existing SciLifeLab Platform(s) would the facility fit": "Fac_fits",
        "Indicate if the suggested technology/instrument/service/technological capability would considerably contribute to strengthen one or more of the SciLifeLab capabilities and/or the Data Driven Life Science program": "cap_fits_A",
        "Indicate if the suggested facility would considerably contribute to strengthen one or more of the SciLifeLab capabilities and/or the Data Driven Life Science program": "cap_fits_B",
        "Estimate the number of unique annual users if the unit would become a part of the SciLifeLab national infrastructure": "potential_users",
    },
    inplace=True,
)

# made where the tech/facility fits in one column (for which platform does it fit in question)

survey_data_raw["Platform_fits"] = (
    survey_data_raw["Tech_fits"] + survey_data_raw["Fac_fits"]
)

# made which capability would be contributed to fit in one column (for which capability does it fit in question)

survey_data_raw["Capability_fits"] = (
    survey_data_raw["cap_fits_A"] + survey_data_raw["cap_fits_B"]
)

# print(survey_data_raw.info())
# There are two different sets of plots needed (one for each survey type), although some plots are needed for both types
# Split data according to survey type (A and B)

surveyA = survey_data_raw[
    (
        survey_data_raw
        == "a.	From a user perspective, an urgently needed technology, instrument, service, or technological capability, currently not available as nation-wide service in Sweden"
    ).any(axis=1)
]

surveyB = survey_data_raw[
    (
        survey_data_raw
        == "b.	An existing local or national core-facility that could be incorporated as a SciLifeLab unit from 2025"
    ).any(axis=1)
]

# Need counts for each survey types (go into top of pages)

countA = surveyA.shape[0]
countB = surveyB.shape[0]

# Below here is all plots and associated data preparation

### Make affiliation plots - needed for both survey types

# Dataframe of all possible values (needed to ensure all values can be on the plot, even if not selected in the survey)

Aff_data = pd.DataFrame(
    {
        "University": [
            "Chalmers University of Technology",
            "Karolinska Institutet",
            "KTH, Royal Institute of Technology",
            "Linköping University",
            "Lund University",
            "Stockholm University",
            "Swedish University of Agricultural Sciences",
            "Umeå University",
            "University of Gothenburg",
            "Uppsala University",
            "Örebro University",
            "Other Swedish University",
            "",
            "",
            "",
            "Other University",
        ],
        "Affiliation": [
            "Chalmers University of Technology",
            "Karolinska Institutet",
            "KTH, Royal Institute of Technology",
            "Linköping University",
            "Lund University",
            "Stockholm University",
            "Swedish University of Agricultural Sciences",
            "Umeå University",
            "University of Gothenburg",
            "Uppsala University",
            "Örebro University",
            "Other Swedish University",
            "Governmental organization",
            "Healthcare",
            "Industry",
            "Other University",
        ],
    }
)

# get counts for affiliations of those that aubmitted survey type A

affiliationsA = surveyA[["Affiliation", "University"]]

affA_comb = pd.concat([Aff_data, affiliationsA])

affiliationsA = (
    affA_comb.groupby(["University", "Affiliation"]).size().reset_index(name="Count")
)
affiliationsA["Count"] = affiliationsA["Count"] - 1

# get counts for affiliations of those that aubmitted survey type B

affiliationsB = surveyB[["Affiliation", "University"]]

affB_comb = pd.concat([Aff_data, affiliationsB])

affiliationsB = (
    affB_comb.groupby(["University", "Affiliation"]).size().reset_index(name="Count")
)
affiliationsB["Count"] = affiliationsB["Count"] - 1

# now make affiliations plot


def affiliations_bar(input, name, colour):
    affiliations = input
    fig = go.Figure(
        data=[
            go.Bar(
                name="Affiliations",
                y=affiliations.Affiliation,
                x=affiliations.Count,
                orientation="h",
                marker=dict(color=colour, line=dict(color="#000000", width=1)),
            ),
        ]
    )

    fig.update_layout(
        plot_bgcolor="white",
        font=dict(size=12),
        autosize=False,
        margin=dict(r=250, t=0, b=0, l=0),
        width=600,
        height=600,
    )

    # modify y-axis
    fig.update_yaxes(
        title=" ",
        showgrid=True,
        linecolor="black",
        ticktext=[
            "<b>Chalmers University of Technology</b>",
            "<b>Karolinska Institutet</b>",
            "<b>KTH, Royal Institute of Technology</b>",
            "<b>Linköping University</b>",
            "<b>Lund University</b>",
            "<b>Stockholm University</b>",
            "<b>Swedish University of Agricultural Sciences</b>",
            "<b>Umeå University</b>",
            "<b>University of Gothenburg</b>",
            "<b>Uppsala University</b>",
            "<b>Örebro University</b>",
            "<b>Other Swedish University</b>",
            "<b>Governmental organization</b>",
            "<b>Healthcare</b>",
            "<b>Industry</b>",
            "<b>Other University</b>",
        ],
        tickvals=[
            "Chalmers University of Technology",
            "Karolinska Institutet",
            "KTH, Royal Institute of Technology",
            "Linköping University",
            "Lund University",
            "Stockholm University",
            "Swedish University of Agricultural Sciences",
            "Umeå University",
            "University of Gothenburg",
            "Uppsala University",
            "Örebro University",
            "Other Swedish University",
            "Governmental organization",
            "Healthcare",
            "Industry",
            "Other University",
        ],
        categoryorder="array",
        categoryarray=[
            "Industry",
            "Healthcare",
            "Governmental organization",
            "Other Swedish University",
            "Other University",
            "Örebro University",
            "Uppsala University",
            "University of Gothenburg",
            "Umeå University",
            "Swedish University of Agricultural Sciences",
            "Stockholm University",
            "Lund University",
            "Linköping University",
            "KTH, Royal Institute of Technology",
            "Karolinska Institutet",
            "Chalmers University of Technology",
        ],
    )

    # modify x-axis
    fig.update_xaxes(
        title=" ",
        showgrid=True,
        gridcolor="black",
        linecolor="black",
        dtick=2,
        range=[0, int(max(affiliations.Count * 1.15))],
    )
    # fig.show()

    if not os.path.isdir("Plots/"):
        os.mkdir("Plots/")
    fig.write_image("Plots/affiliation_{}.svg".format(name))


# function to iterate through

affiliations_bar(affiliationsA, "A", "#4C979F")
affiliations_bar(affiliationsB, "B", "#A7C947")

### In which Platform would it fit? - for both survey types, although slight difference in exactly what's recorded for each type

# We need to use the Platform_fits column, but since can have multiple units listed in that column, it's necessary to do the counts as substrings

# work to make sure that zero values (i.e. survey options not selected are included)
Plat_data = pd.DataFrame(
    {
        "Platform": [
            "Genomics",
            "Clinical Genomics",
            "Metabolomics",
            "Spatial Biology",
            "Cellular and Molecular Imaging",
            "Integrated Structural Biology",
            "Chemical Biology and Genome Engineering",
            "Clinical Proteomics and Immunology",
            "Drug Discovery and Development",
            "Bioinformatics",
            "None of the existing platforms",
            "I do not know",
        ],
        "Count": [
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
            0,
        ],
    }
)

Plat_fit_A = pd.DataFrame(
    surveyA.Platform_fits.str.extractall("({})".format("|".join(Plat_data["Platform"])))
    .iloc[:, 0]
    .str.get_dummies()
    .sum()
    .reset_index()
    .rename(columns={"index": "Platform", 0: "Count"})
)

plata_comb = pd.concat([Plat_data, Plat_fit_A])

plata = plata_comb.groupby(["Platform"]).sum().reset_index()

Plat_fit_B = (
    surveyB.Platform_fits.str.extractall("({})".format("|".join(Plat_data["Platform"])))
    .iloc[:, 0]
    .str.get_dummies()
    .sum()
    .reset_index()
    .rename(columns={"index": "Platform", 0: "Count"})
)

platb_comb = pd.concat([Plat_data, Plat_fit_B])

platb = platb_comb.groupby(["Platform"]).sum().reset_index()


# plot
def platform_fit_bar(input, name, colour):
    plat_fit = input
    fig = go.Figure(
        data=[
            go.Bar(
                name="Platform",
                y=plat_fit.Platform,
                x=plat_fit.Count,
                orientation="h",
                marker=dict(color=colour, line=dict(color="#000000", width=1)),
            ),
        ]
    )

    fig.update_layout(
        plot_bgcolor="white",
        font=dict(size=12),
        autosize=False,
        margin=dict(r=250, t=0, b=0, l=0),
        width=600,
        height=600,
    )

    # modify y-axis
    fig.update_yaxes(
        title=" ",
        showgrid=True,
        linecolor="black",
        ticktext=[
            "<b>I do not know</b>",
            "<b>None of the existing platforms</b>",
            "<b>Spatial Biology</b>",
            "<b>Metabolomics</b>",
            "<b>Integrated Structural Biology</b>",
            "<b>Genomics</b>",
            "<b>Drug Discovery and Development</b>",
            "<b>Clinical Proteomics and Immunology</b>",
            "<b>Clinical Genomics</b>",
            "<b>Chemical Biology and Genome Engineering</b>",
            "<b>Cellular and Molecular Imaging</b>",
            "<b>Bioinformatics</b>",
        ],
        tickvals=[
            "I do not know",
            "None of the existing platforms",
            "Spatial Biology",
            "Metabolomics",
            "Integrated Structural Biology",
            "Genomics",
            "Drug Discovery and Development",
            "Clinical Proteomics and Immunology",
            "Clinical Genomics",
            "Chemical Biology and Genome Engineering",
            "Cellular and Molecular Imaging",
            "Bioinformatics",
        ],
        categoryorder="array",
        categoryarray=[
            "I do not know",
            "None of the existing platforms",
            "Spatial Biology",
            "Metabolomics",
            "Integrated Structural Biology",
            "Genomics",
            "Drug Discovery and Development",
            "Clinical Proteomics and Immunology",
            "Clinical Genomics",
            "Chemical Biology and Genome Engineering",
            "Cellular and Molecular Imaging",
            "Bioinformatics",
        ],
    )

    # modify x-axis
    fig.update_xaxes(
        title=" ",
        showgrid=True,
        gridcolor="black",
        linecolor="black",
        dtick=2,
        range=[0, int(max(plat_fit.Count * 1.15))],
    )
    # fig.show()

    if not os.path.isdir("Plots/"):
        os.mkdir("Plots/")
    fig.write_image("Plots/platform_fit_{}.svg".format(name))


# function to iterate through

platform_fit_bar(plata, "A", "#4C979F")
platform_fit_bar(platb, "B", "#A7C947")

### Contribution to capabilities - needed for both survey types

Capability_data = pd.DataFrame(
    {
        "Capability": [
            "Pandemic Laboratory Preparedness",
            "Precision Medicine",
            "Planetary Biology",
            "Data Driven Life Science",
            "None",
            "I do not know",
        ],
        "Count": [
            0,
            0,
            0,
            0,
            0,
            0,
        ],
    }
)

Capability_fit_A = pd.DataFrame(
    surveyA.Capability_fits.str.extractall(
        "({})".format("|".join(Capability_data["Capability"]))
    )
    .iloc[:, 0]
    .str.get_dummies()
    .sum()
    .reset_index()
    .rename(columns={"index": "Capability", 0: "Count"})
)

capa_comb = pd.concat([Capability_data, Capability_fit_A])

capa = capa_comb.groupby(["Capability"]).sum().reset_index()

Capability_fit_B = pd.DataFrame(
    surveyB.Capability_fits.str.extractall(
        "({})".format("|".join(Capability_data["Capability"]))
    )
    .iloc[:, 0]
    .str.get_dummies()
    .sum()
    .reset_index()
    .rename(columns={"index": "Capability", 0: "Count"})
)

capb_comb = pd.concat([Capability_data, Capability_fit_B])

capb = capb_comb.groupby(["Capability"]).sum().reset_index()


# Plot
def capability_fit_bar(input, name, colour):
    capability_fit = input
    fig = go.Figure(
        data=[
            go.Bar(
                name="Capability",
                y=capability_fit.Capability,
                x=capability_fit.Count,
                orientation="h",
                marker=dict(color=colour, line=dict(color="#000000", width=1)),
            ),
        ]
    )

    fig.update_layout(
        plot_bgcolor="white",
        font=dict(size=12),
        autosize=False,
        margin=dict(r=250, t=0, b=0, l=0),
        width=600,
        height=600,
    )

    # modify y-axis
    fig.update_yaxes(
        title=" ",
        showgrid=True,
        linecolor="black",
        ticktext=[
            "<b>Pandemic Laboratory Preparedness</b>",
            "<b>Precision Medicine</b>",
            "<b>Planetary Biology</b>",
            "<b>Data Driven Life Science</b>",
            "<b>None</b>",
            "<b>I do not know</b>",
        ],
        tickvals=[
            "Pandemic Laboratory Preparedness",
            "Precision Medicine",
            "Planetary Biology",
            "Data Driven Life Science",
            "None",
            "I do not know",
        ],
        categoryorder="array",
        categoryarray=[
            "I do not know",
            "None",
            "Precision Medicine",
            "Planetary Biology",
            "Pandemic Laboratory Preparedness",
            "Data Driven Life Science",
        ],
    )

    # modify x-axis
    fig.update_xaxes(
        title=" ",
        showgrid=True,
        gridcolor="black",
        linecolor="black",
        dtick=2,
        range=[0, int(max(capability_fit.Count * 1.15))],
    )
    # fig.show()

    if not os.path.isdir("Plots/"):
        os.mkdir("Plots/")
    fig.write_image("Plots/capability_fit_{}.svg".format(name))


# # function to iterate through

capability_fit_bar(capa, "A", "#4C979F")
capability_fit_bar(capb, "B", "#A7C947")

# Estimate number of users that would have if incorporated into SciLifeLab - only needed for survey type B
# Can only select one option here, so no need to split strings.

# working to ensure that even options not selected in the survey are included
Potential_users_data = pd.DataFrame(
    {
        "potential_users": [
            "1-10",
            "10-50",
            "More than 50",
            "I do not know",
        ],
        "Count": [
            0,
            0,
            0,
            0,
        ],
    }
)

potential_users_counts = (
    surveyB.groupby(["potential_users"]).size().reset_index(name="Count")
)

pot_users_comb = pd.concat([Potential_users_data, potential_users_counts])

pot_users = pot_users_comb.groupby(["potential_users"]).sum().reset_index()


# plot
def potential_users_bar(input, name, colour):
    pot_users = input
    fig = go.Figure(
        data=[
            go.Bar(
                name="Potential Users",
                y=pot_users.potential_users,
                x=pot_users.Count,
                orientation="h",
                marker=dict(color=colour, line=dict(color="#000000", width=1)),
            ),
        ]
    )

    fig.update_layout(
        plot_bgcolor="white",
        font=dict(size=12),
        autosize=False,
        margin=dict(r=250, t=0, b=0, l=0),
        width=600,
        height=600,
    )

    # modify y-axis
    fig.update_yaxes(
        title=" ",
        showgrid=True,
        linecolor="black",
        ticktext=[
            "<b>1-10<b>",
            "<b>10-50",
            "<b>More than 50",
            "<b>I do not know",
        ],
        tickvals=[
            "1-10",
            "10-50",
            "More than 50",
            "I do not know",
        ],
        categoryorder="array",
        categoryarray=[
            "I do not know",
            "More than 50",
            "10-50",
            "1-10",
        ],
    )

    # modify x-axis
    fig.update_xaxes(
        title=" ",
        showgrid=True,
        gridcolor="black",
        linecolor="black",
        dtick=2,
        range=[0, int(max(pot_users.Count * 1.15))],
    )
    # fig.show()

    if not os.path.isdir("Plots/"):
        os.mkdir("Plots/")
    fig.write_image("Plots/potential_users_{}.svg".format(name))


# function to iterate through
potential_users_bar(pot_users, "B", "#A7C947")
