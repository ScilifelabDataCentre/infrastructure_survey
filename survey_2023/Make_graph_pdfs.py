# Regular packages
import os

import pandas as pd

# Specific imports from reportlab
from reportlab.platypus import (
    BaseDocTemplate,
    Paragraph,
    Spacer,
    Image,
    PageTemplate,
    Frame,
    CondPageBreak,
)
from reportlab.platypus.flowables import HRFlowable
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Arial fonts
pdfmetrics.registerFont(TTFont("Arial", "Arial.ttf"))
pdfmetrics.registerFont(TTFont("Arial-B", "Arial Bold.ttf"))
pdfmetrics.registerFont(TTFont("Arial-I", "Arial Italic.ttf"))
pdfmetrics.registerFont(TTFont("Arial-N", "Arial Narrow.ttf"))

# This import facilitates the header creation
from functools import partial

# SVG file function
from svglib.svglib import svg2rlg

# These is the data to import for these pages
from Make_plots import countA, countB


def header(canvas, doc, content):
    """
    header creates a header for a reportlabs document, and is inserted in the template
    """
    canvas.saveState()
    w, h = content.wrap(doc.width, doc.topMargin)
    content.drawOn(canvas, doc.leftMargin, doc.height + doc.topMargin - h)
    p = canvas.beginPath()
    p.moveTo(doc.leftMargin, doc.height + doc.topMargin - h - 2 * mm)
    p.lineTo(doc.leftMargin + w, doc.height + doc.topMargin - h - 2 * mm)
    p.close()
    canvas.setLineWidth(0.5)
    canvas.setStrokeColor("#A6A6A6")
    canvas.drawPath(p, stroke=1)
    canvas.restoreState()


def generatePdf(
    Survey_name,
):  # going to make two types of page (one for survey type A and one for survey type B,)
    """
    generatePdf creates a PDF document based on the reporting data supplied.
    It is using very strict formatting, but is quite simple to edit.
    This function will print the name of the unit its working on, and
    any warnings that may arise. The excel document can be edited to fix warnings
    and to change the information in the PDFs.
    """
    if not os.path.isdir("pdfs_plots/"):
        os.mkdir("pdfs_plots/")
    # Setting the document sizes and margins. showBoundary is useful for debugging
    doc = BaseDocTemplate(
        "pdfs_plots/plots_survey{}.pdf".format(Survey_name.lower()),
        pagesize=A4,
        rightMargin=18 * mm,
        leftMargin=14 * mm,
        topMargin=16 * mm,
        bottomMargin=20 * mm,
        showBoundary=0,
    )
    # These are the fonts available, in addition to a number of "standard" fonts.
    # These are used in setting paragraph styles
    # I have used spaceAfter, spaceBefore and leading to change the layout of the "paragraphs" created with these styles
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="onepager_inner_heading",  # This is what we should use for the counts
            parent=styles["Heading1"],
            fontName="Arial-B",
            fontSize=13,  # consider perhaps a fontsize a little bigger (but less than 16, as that's the header)
            color="#FF00AA",
            leading=16,
            spaceAfter=0,
            spaceBefore=8,
        )
    )
    styles.add(
        ParagraphStyle(
            name="chart_heading",  # use this for all chart headers
            parent=styles["Heading1"],
            fontName="Arial-B",
            fontSize=10,
            color="#FF00AA",
            leading=16,
            spaceAfter=0,
            spaceBefore=8,
        )
    )
    styles.add(
        ParagraphStyle(
            name="onepager_title",  # This is the header title
            parent=styles["Heading1"],
            fontName="Arial-B",
            fontSize=16,
            bold=0,
            color="#000000",
            leading=16,
            spaceBefore=0,
        )
    )
    #     styles.add(
    #         ParagraphStyle(
    #             name="onepager_text",
    #             parent=styles["Normal"],
    #             fontName="Arial",
    #             fontSize=10,
    #             bold=0,
    #             color="#000000",
    #             leading=14,
    #         )
    #     )
    #     styles.add(
    #         ParagraphStyle(
    #             name="onepager_footnote",
    #             parent=styles["Normal"],
    #             fontName="Lato",
    #             fontSize=7,
    #             bold=0,
    #             color="#000000",
    #             leading=14,
    #             spaceBefore=20,
    #         )
    #     )
    # The document is set up with frames, each frame incorporates part of the page
    # The first frame includes information about the total number of proposals
    frame1 = Frame(
        doc.leftMargin,
        doc.bottomMargin + (doc.height / 2),
        doc.width / 2,  # - 3.5 * mm,
        (doc.height / 2) - 13 * mm,
        id="col1",
        # showBoundary=1,  # this can be used to show where the frame sits - useful for setting up layout
        leftPadding=0 * mm,
        topPadding=0 * mm,
        rightPadding=0 * mm,
        bottomPadding=0 * mm,
    )
    # The second frame deals with the first column for graphs
    frame2 = Frame(
        doc.leftMargin,  # determines how left on the page it sits
        doc.bottomMargin,  # determines how high on page it sits
        doc.width / 2,  # determines how wide it is
        (doc.height) - 20 * mm,  # determines how tall it is
        id="pic1",
        # showBoundary=1, # Useful for testing where the box is
        # all below just gives a little padding
        leftPadding=3 * mm,
        topPadding=3 * mm,
        rightPadding=0 * mm,
        bottomPadding=0 * mm,
    )
    # The third frame deals with the second column of graphs
    frame3 = Frame(
        doc.leftMargin + (doc.width / 2),
        doc.bottomMargin,  # + (doc.height / 2),
        doc.width / 2,  # 2 - 3.5 * mm,
        (doc.height) - 20 * mm,  # - 20 * mm,  # + 50 * mm,  # - 18 * mm,
        id="pic2",
        # showBoundary=1,
        leftPadding=3 * mm,
        topPadding=3 * mm,
        rightPadding=0 * mm,
        bottomPadding=0 * mm,
    )

    pd.options.display.max_colwidth = 600
    if Survey_name == "A":
        header_content = Paragraph(
            "<font color='#4C979F' name=Arial-B><b>Proposals on New Technologies – a summary</b></font>",
            styles["onepager_title"],
        )
    else:
        header_content = Paragraph(
            "<font color='#A7C947' name=Arial-B><b>Proposals on New Infrastructure Units – a summary</b></font>",
            styles["onepager_title"],
        )
    template = PageTemplate(
        id="test",
        frames=[frame1, frame2, frame3],
        onPage=partial(header, content=header_content),
    )
    doc.addPageTemplates([template])
    # The Story list will contain all Paragraph and other elements. In the end this is used to build the document
    Story = []
    ### Below here will be Paragraph and Image elements added to the Story, they flow through frames automatically,
    ### however I have set a framebreak to correctly organise things in left/right column.
    pd.options.display.max_colwidth = 600
    if Survey_name == "A":
        Story.append(
            Paragraph(
                "<font color='#4C979F' name=Arial-B><b>Total number of proposals: {}</b></font>".format(
                    (countA),  # .to_string(index=False),
                ),
                styles["onepager_inner_heading"],
            )
        )
    else:
        Story.append(
            Paragraph(
                "<font color='#A7C947' name=Arial-B><b>Total number of proposals: {}</b></font>".format(
                    (countB),  # .to_string(index=False),
                ),
                styles["onepager_inner_heading"],
            )
        )
    # Now to put in figures
    # (for both: Affiliation of proposer, in which platforms would it fit, and in which capability would it fit)
    # (for just survey B: Estimate number of potential users)
    # figs already made in .svg format, they need to be imported
    Story.append(CondPageBreak(200 * mm))  # move to next frame
    # First put affiliations plot in
    if Survey_name == "A":
        Story.append(
            Paragraph(
                "<font color='#4C979F' name=Arial-B><b>Affiliation of Proposer:</b></font>",
                styles["chart_heading"],
            )
        )
        filepath_affiliation = "Plots/affiliation_{}.svg".format(
            (Survey_name),
        )
        isFile_affiliation = os.path.isfile(filepath_affiliation)
        im_affiliation = svg2rlg(filepath_affiliation)
        im_affiliation = Image(im_affiliation, width=80 * mm, height=60 * mm)
        im_affiliation.hAlign = "LEFT"
        Story.append(im_affiliation)
    else:
        Story.append(
            Paragraph(
                "<font color='#A7C947' name=Arial-B><b>Affiliation of Proposer:</b></font>",
                styles["chart_heading"],
            )
        )
        filepath_affiliation = "Plots/affiliation_{}.svg".format(
            (Survey_name),
        )
        isFile_affiliation = os.path.isfile(filepath_affiliation)
        im_affiliation = svg2rlg(filepath_affiliation)
        im_affiliation = Image(im_affiliation, width=80 * mm, height=60 * mm)
        im_affiliation.hAlign = "LEFT"
        Story.append(im_affiliation)
    # Next, plot related to which platform it fits into
    # The title for the question needs to be slightly different
    # Story.append(CondPageBreak(10 * mm))
    # "<font color='#4C979F' name=Arial-B><b></b></font>",
    #             "<font color='#A7C947' name=Arial-B><b></b></font>",
    if Survey_name == "A":
        Story.append(
            Paragraph(
                "<font color='#4C979F' name=Arial-B><b>In which SciLifeLab Platform would the technology/instrument/service/technological capability fit?</b></font>",
                styles["chart_heading"],
            )
        )
        filepath_platform = "Plots/platform_fit_{}.svg".format(
            (Survey_name),
        )
        isFile_platform = os.path.isfile(filepath_platform)
        im_platform = svg2rlg(filepath_platform)
        im_platform = Image(im_platform, width=80 * mm, height=60 * mm)
        im_platform.hAlign = "LEFT"
        Story.append(im_platform)
    else:
        Story.append(
            Paragraph(
                "<font color='#A7C947' name=Arial-B><b>In which SciLifeLab Platform would the facility/unit fit?</b></font>",
                styles["chart_heading"],
            )
        )
        filepath_platform = "Plots/platform_fit_{}.svg".format(
            (Survey_name),
        )
        isFile_platform = os.path.isfile(filepath_platform)
        im_platform = svg2rlg(filepath_platform)
        im_platform = Image(im_platform, width=80 * mm, height=60 * mm)
        im_platform.hAlign = "LEFT"
        Story.append(im_platform)
    # Next, plot related to which capability/program could be strengthened
    # The title for the question needs to be slightly different
    if Survey_name == "A":
        Story.append(
            Paragraph(
                "<font color='#4C979F' name=Arial-B><b>Which SciLifeLab capability/program could potentially be strengthened by the technology/instrument/service/technological capability?</b></font>",
                styles["chart_heading"],
            )
        )
        filepath_capability = "Plots/capability_fit_{}.svg".format(
            (Survey_name),
        )
        isFile_capability = os.path.isfile(filepath_capability)
        im_capability = svg2rlg(filepath_capability)
        im_capability = Image(im_capability, width=80 * mm, height=60 * mm)
        im_capability.hAlign = "LEFT"
        Story.append(im_capability)
    else:
        Story.append(
            Paragraph(
                "<font color='#A7C947' name=Arial-B><b>Which SciLifeLab capability/program could potentially be strengthened by the facility/unit?</b></font>",
                styles["chart_heading"],
            )
        )
        filepath_capability = "Plots/capability_fit_{}.svg".format(
            (Survey_name),
        )
        isFile_capability = os.path.isfile(filepath_capability)
        im_capability = svg2rlg(filepath_capability)
        im_capability = Image(im_capability, width=80 * mm, height=60 * mm)
        im_capability.hAlign = "LEFT"
        Story.append(im_capability)
        # now need last graph for survey type B (new frame needed) - Estimate number of potential users
        # No further graphs for survey type A
        Story.append(CondPageBreak(200 * mm))  # move to next frame
        Story.append(
            Paragraph(
                "<font color='#A7C947' name=Arial-B><b>Estimated number of unique visitors annually if the facility/unit became part of SciLifeLab's national infrastructure:</b></font>",
                styles["chart_heading"],
            )
        )
        filepath_potuse = "Plots/potential_users_B.svg"
        isFile_potuse = os.path.isfile(filepath_potuse)
        im_potuse = svg2rlg(filepath_potuse)
        im_potuse = Image(im_potuse, width=80 * mm, height=60 * mm)
        im_potuse.hAlign = "LEFT"
        Story.append(im_potuse)

    #
    # if Survey_name == "A":
    #     # Story.append(
    #     #     Paragraph(
    #     #         "<font color='#4C979F' name=Arial-B><b>In which SciLifeLab Platform would the technology/instrument/service/technological capability fit?</b></font>",
    #     #         styles["chart_heading"],
    #     #     )
    #     # )
    #     filepath_platform = "Plots/platform_{}.svg".format(
    #         (Survey_name),
    #     )
    #     isFile_platform = os.path.isfile(filepath_platform)
    #     im_platform = svg2rlg(filepath_platform)
    #     im_platform = Image(im_platform, width=70 * mm, height=55 * mm)
    #     im_platform.hAlign = "CENTER"
    #     Story.append(im_platform)
    # else:
    #     Story.append(
    #         Paragraph(

    #             styles["chart_heading"],
    #         )
    #     )
    #     filepath_platform = "Plots/platform_{}.svg".format(
    #         (Survey_name),
    #     )
    #     isFile_platform = os.path.isfile(filepath_platform)
    #     im_platform = svg2rlg(filepath_platform)
    #     im_platform = Image(im_platform, width=70 * mm, height=55 * mm)
    #     im_platform.hAlign = "CENTER"
    #     Story.append(im_platform)
    # if Survey_name == "A":
    #     Story.append(
    #     Paragraph(
    #         "<font color='#A7C947' name=Lato-B><b></b></font>",
    #         styles["onepager_chart_heading"],
    #     )
    # ),
    # filepath_platform = "Plots/platform_{}.svg".format(
    #     (Survey_name),
    # )
    # else:
    #     Story.append(
    #     Paragraph(
    #         "<font color='#A7C947' name=Lato-B><b></b></font>",
    #         styles["onepager_chart_heading"],
    #     )
    # ),
    # filepath_platform = "Plots/platform_{}.svg".format(
    #     (Survey_name),
    # )

    # Finally, build the document.
    doc.build(Story)


# Note: not setting the year universally, because it might be that you're reporting for the current year, or the one before
generatePdf("A")
generatePdf("B")
