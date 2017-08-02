# -*- coding: utf-8 -*-
import sys
from utils import *
from sheets_api import *
from report import Report
from village import Village
from missionary import Missionary, Child, Spouse
import logging
from imgur import update_imgur_ids, get_image
from powerpoint import *
from Queue import Queue
import threading
import argparse

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
    logging.basicConfig(filename='diags.log',
                        level=logging.INFO,
                        format='%(asctime)s %(name)-12s %(levelname)-8s %(threadName)s %(message)s',
                        datefmt='%m-%d %H:%M')
    parser = argparse.ArgumentParser(description='Report builder for 500k')
    parser.add_argument('-d',
                        '--date',
                        help='Month and year to generate reports for in MM/YYYY integer form e.g. 01/2017',
                        required=False)
    parser.add_argument('--test', help='Use sample data only', action='store_true')
    args = parser.parse_args()

    # Validate date input
    if args.date:
        try:
            if len(args.date) != 7:
                raise TypeError
            int(args.date[:2])
            int(args.date[3:])
        except TypeError:
            raise ValueError("Date must be in MM/YYYY format")

    # Make sure all Imgur IDs are up-to-date.
    imgur_imgs = update_imgur_ids()

    # Gather all information from the spreadsheet. Returned as list of lists
    # where each list is a row of cells.
    if args.test:
        report_data = get_all_missionary_reports(test=True)
        factfile_data = get_all_factfile_data(test=True)
    else:
        report_data = get_all_missionary_reports()
        factfile_data = get_all_factfile_data()

    # Now build out the data into usable dictionaries
    all_missionaries = construct_data(report_data, factfile_data, imgur_imgs)

    # Time to create the presentations, loop around for every single missionary
    # TODO - In future make sure only missionaries with new reports get
    # generated
    logging.info("Creating powerpoints for {0} missionaries".format(
        len(all_missionaries)))

    q = Queue(maxsize=0)
    num_threads = 10

    for i in range(num_threads):
        worker = threading.Thread(target=create_powerpoint_pdf, args=(q,))
        worker.setDaemon(True)
        worker.start()

    if args.date:
        date = args.date
    else:
        date = None

    for miss_id, missionary in all_missionaries.iteritems():
        q.put((missionary, miss_id, date))

    q.join()

    return 0


def construct_data(report_data, factfile_data, imgur_imgs):
    """
    Take from the two different spreadsheets to create a total view of all the
    missionary data, once complete we have all the info required to start
    creating the reports.
    """
    all_missionaries = {}
    construct_factfile_data(all_missionaries, factfile_data)
    construct_report_data(all_missionaries, report_data)
    # add_imgur_profiles(all_missionaries, imgur_imgs)
    return all_missionaries


def construct_report_data(all_missionaries, report_data):
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
    #          -> All Prayer Points
    #           ...
    #      -> Report 2
    #       ...
    #  -> Missionary 2
    #   ...
    logging.info("Constructing report data")

    # As we may change the order of the columns from time to time and need to
    # make this sustainable for any changes, create a dictionary of column
    # numbers against the column header text. This should be a single linear
    # search to get all the headings.
    columns = dict()
    for idx, column in enumerate(report_data[0]):
        columns[column] = idx

    for row in report_data[1:]:
        if len(row) > columns[u'\u2022Main Story / Report: ']:
            try:
                report = Report(row[columns['Date (Pretty)']],
                                row[columns[u'\u2022Missionary Name: ']],
                                row[columns[u'\u2022MissionaryID: ']],
                                row[columns[u'\u2022Main Story / Report: ']])
            except NotImplementedError:
                continue
            villages = []
            prayer_rqs = []
            for i in range(1, 6):
                if row[columns[u'\u2022V' + str(i) + ': ']]:
                    villages.append(
                        Village(row[columns[u'\u2022V' + str(i) + ': ']],
                                row[columns[u'\u2022V' + str(i) + 'N: ']],
                                row[columns[u'\u2022V' + str(i) + 'B: ']]))
            for i in range(1, 9):
                if (len(row) > columns["P-R-" + str(i) + ": "] and
                        row[columns["P-R-" + str(i) + ": "]]):
                    prayer_rqs.append(row[columns["P-R-" + str(i) + ": "]])
            report.villages = villages
            report.prayer_rqs = prayer_rqs
            report.round = report.get_report_round()
            missionary_id = report.id
            if missionary_id in all_missionaries.keys():
                # Missionary already exists, add report to dictionary
                missionary = all_missionaries[missionary_id]
            else:
                # New missionary, create new missionary and add report.
                logging.warning("No factfile data for {0}".format(
                    missionary_id))
                names = report.name.split(" ")
                if len(names) > 1:
                    try:
                        missionary = Missionary(missionary_id,
                                                names[-1],
                                                names[-2])
                    except NotImplementedError:
                        continue
                else:
                    try:
                        missionary = Missionary(missionary_id,
                                                names[-1],
                                                None)
                    except NotImplementedError:
                        continue
                all_missionaries[missionary_id] = missionary
            missionary.reports[report.round] = report

    logging.info("Report data has been constructed")


def construct_factfile_data(all_missionaries, factfile_data):
    """
    Start building the Missionary data using factfile information
    """
    logging.info("Constructing factfile data")

    # As we may change the order of the columns from time to time and need to
    # make this sustainable for any changes, create a dictionary of column
    # numbers against the column header text. This should be a single linear
    # search to get all the headings.
    columns = dict()
    for idx, column in enumerate(factfile_data[0]):
        columns[column] = idx

    for row in factfile_data[1:]:
        if len(row) > columns[u'Profile Picture']:
            # Basics mandatory for a factfile
            try:
                missionary = Missionary(row[columns[u'ID (new)']],
                                        row[columns[u'MissionarySecondName']],
                                        row[columns[u'MissionaryFirstName']])
            except NotImplementedError:
                logging.error("Couldn't create {0} factfile data".format(
                    row[columns[u'ID (new)']]))
            try:
                missionary.state = validate_state(
                    row[columns[u'MissionField State']])
            except ValueError:
                logging.error("Invalid state for {0}: {1}".format(
                    row[columns[u'ID (new)']],
                    row[columns[u'MissionField State']]))
            missionary.pic = row[columns[u'Headshot Photo link']]
            # Add family and biography
            if len(row) > columns[u'Number of Dependents']:
                if row[columns[u'Wife / Husband\'s First Name']]:
                    missionary.spouse = Spouse(
                        row[columns[u'Wife / Husband\'s First Name']],
                        row[columns[u'Wife / Husband\'s Second Name']])
            for i in range(1, 6):
                if row[columns[u'Child ' + str(i) + ' First Name']]:
                    missionary.children[u'Child ' + str(i)] = Child(
                        row[columns[u'Child ' + str(i) + ' First Name']],
                        row[columns[u'Child ' + str(i) + ' DOB']])

            # Mission field data
            villages = []
            prayer_rqs = []
            for i in range(1, 6):
                if (len(row) > columns[u'V' + str(i) + ' B'] and
                        row[columns[u'V' + str(i)]]):
                    villages.append(
                        Village(row[columns[u'V' + str(i)]],
                                row[columns[u'V' + str(i) + ' N']],
                                row[columns[u'V' + str(i) + ' B']]))
            missionary.villages = villages
            all_missionaries[missionary.id] = missionary

    logging.info("Factfile data has been constructed")


def add_imgur_profiles(all_missionaries, imgur_imgs):
    for miss_id, missionary in all_missionaries.iteritems():
        try:
            missionary.pic = imgur_imgs[miss_id]
        except KeyError:
            logging.info('{0} has no Imgur picture'.format(miss_id))


if __name__ == '__main__':
    status = main()
    sys.exit(status)
