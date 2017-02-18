from pptx import Presentation
from docx import Document
from imgurpython import ImgurClient
from datetime import datetime
import sys, os, shutil
from utils import *
from sheets_api import *

# This script conducts the following:
# - Gets the information on a missionary based on Miss ID (gets all)
# - Populates the title slide with the missionaries details
# - Creates a content slide per report
# - Exports the slideshow to pdf
# - Uploads to Ben's Google Drive account via API
# - Pastes the URL of the report in Google Drive to the web report sheets
#
# With this info Ben can then send out multiple reports using:
# https://support.yet-another-mail-merge.com/hc/en-us/articles/210735349

def main(argv=None):
    # Gather all information from the spreadsheet. Returned as list of lists
    # where each list is a row of cells.
    values = get_all_missionary_reports()

    all_dict = construct_data(values)

    # Time to create the presentations, loop around for every single missionary
    # TODO - In future make sure only missionaries with new reports get
    # generated
    for miss_id,miss_dict in all_dict:
        pptx = create_powerpoint(miss_id,miss_dict)

        # Export to pdf
        PPTtoPDF(pptx, pptx.split(".")[0] + ".pdf")

def construct_data(values):
    # For all the missionaries, arrange data in this structure:
    # All
    #  -> Missionary 1 (based on ID)
    #      -> Report 1
    #          -> Date, Subject, Raw, Missionary, Missionary ID etc...
    #          -> Village 1
    #              -> Village
    #              -> People
    #              -> Baptisms
    #          -> Village 2
    #           ...
    #          -> Prayer Point 1
    #          -> Prayer Point 2
    #           ...
    #      -> Report 2
    #       ...
    #  -> Missionary 2
    #   ...
    all_dict = dict()
    for row in values:
        report = {"Date":         row[0],
                  "Subject":      row[1],
                  "Raw":          row[3],
                  "Submitter":    row[4],
                  "Email":        row[5],
                  "Missionary":   row[6],
                  "Missionary ID":row[7],
                  "Report":       row[40],
                 }
        for i,village in enumerate(row[8:3:26]):
            if not village.isempty():
                vill_dict = {
                    "Village": row[i+6],
                    "People":  row[i+7],
                    "Baptisms":row[i+8],
                }
                report['Village '+str(i+1)] = vill_dict
        for i,prayer in enumerate(row[41:48]):
            if not prayer.isempty():
                report['Prayer '+str(i+1)] = prayer
        if report["Missionary ID"] in all_dict.keys():
            # Missionary already exists, add report to missionary dictionary
            miss_dict = all_dict[report["Missionary ID"]]
            miss_dict[report["Date"]] = report
        else:
            # New missionary, create new dictionary and add report to it.
            all_dict[report["Missionary ID"]] = {
                report["Missionary ID"]: report}

def create_powerpoint(miss_id,miss_dict):
    # Import presentation
    path = "C:\Users\\br1\Dropbox\NCM\Reports, bills and Proposals\Ben Report Automation\\"
    prs = Presentation(path + "Master Report Template.pptx")

    # Access placeholders for Title slide
    title_slide = prs.slides[0]

    for shape in title_slide.placeholders:
        print('%d %s' % (shape.placeholder_format.idx, shape.name))

    # Insert Missionary Name
    name_holder = title_slide.placeholders[0]
    assert name_holder.has_text_frame
    name_holder.text_frame.clear()
    p = name_holder.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = miss_dict["Missionary"]

    # Insert Missionary ID
    miss_id_holder = title_slide.placeholders[11]
    assert miss_id_holder.has_text_frame
    miss_id_holder.text_frame.clear()
    p = miss_id_holder.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = miss_id

    for k,report in miss_dict:
        # Access placeholders for content slides
        content_slide = prs.slides[1]

        insert_bio(content_slide, report)

        # Report title
        title_holder = content_slide.placeholders[0]
        assert title_holder.has_text_frame
        title_holder.text_frame.clear()
        p = title_holder.text_frame.paragraphs[0]
        run = p.add_run()
        run.text = "<Year> Report <round>"

        # Actual report!
        report_holder = content_slide.placeholders[14]
        assert report_holder.has_text_frame
        report_holder.text_frame.clear()
        p = report_holder.text_frame.paragraphs[0]
        run = p.add_run()
        run.text = report

        # Prayer heading
        prayer_h_holder = content_slide.placeholders[15]
        assert prayer_h_holder.has_text_frame
        prayer_h_holder.text_frame.clear()
        p = prayer_h_holder.text_frame.paragraphs[0]
        run = p.add_run()
        run.text = "Prayer Points"

        # Prayer points body
        prayer_b_holder = content_slide.placeholders[16]
        assert prayer_b_holder.has_text_frame
        prayer_b_holder.text_frame.clear()
        p = prayer_b_holder.text_frame.paragraphs[0]
        run = p.add_run()
        run.text = prayer

    # Save the powerpoint
    print "Where should this report be saved? (Will use default staging area if none.)"
    save_path = raw_input()
    save_name = docx_path.split("\\")[-1].split(".")[0]
    prs.save(save_path + save_name + ".pptx")

def insert_bio(slide, report):
    # Get totals for churches, baptisms and prayer
    for k, value in report_dict:
        # Only villages are nested dictionaries, so test on value is dictionary
        if isinstance(value, dict):
            pass

    # Insert the name and the state of the Missionary
    state_holder = content_slide.placeholders[2]
    assert state_holder.has_text_frame
    state_holder.text_frame.clear()
    p = state_holder.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = "State: " + state_dict[report["Missionary ID"][:2]]

    name_holder = content_slide.placeholders[13]
    assert name_holder.has_text_frame
    name_holder.text_frame.clear()
    p = name_holder.text_frame.paragraphs[0]
    run = p.add_run()
    run.text = report["Missionary"]

    bio_holder = content_slide.placeholders[11]
    assert bio_holder.has_text_frame
    bio_holder.text_frame.clear()
    p = bio_holder.text_frame.paragraphs[0]
    bio_line("\n Churches: ", str(churches), p)
    bio_line("\n Coming for Prayer: ", str(prayer_nos), p)
    bio_line("\n Baptisms: ", str(baptisms), p)

    profile_pic_holder = content_slide.placeholders[10]

    # Insert profile picture, need to access Imgur database!
    profile_pic_holder.insert_picture(img_path)

    # Insert India Map based off state name
    india_pic_holder = content_slide.placeholders[12]
    india_pic_holder.insert_picture('C:\Users\\br1\Dropbox\NCM\Reports, ' +
        'bills and Proposals\!Reporting Workflow\Map Images\\' +
        state_dict[state_ab] + '.png')

if __name__ == '__main__':
    status = main()
    sys.exit(status)
