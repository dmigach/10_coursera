import argparse
import json
import os
import re
import textwrap

from bs4 import BeautifulSoup
from openpyxl import Workbook
import requests

coursera_xml_feed_url = 'https://www.coursera.org/sitemap~www~courses.xml'

search_classes = {
    'course_name': 'title display-3-text',
    'course_language': 'language-info',
    'weeks_tag': 'rc-WeekView',
    'week_tag': 'week'
}


def get_soup_from_url(url):
    """
    :param url: url for html request
    :return: Beautiful Soup object
    """
    request = requests.get(url)
    request.encoding = 'utf-8'
    if request.status_code == 200:
        return BeautifulSoup(request.text, 'lxml')


def get_courses_list(url, amount):
    """
    :param url: url for xml feed with courses list
    :param amount: amount of desired courses
    :return: list with courses urls
    """
    courses_list_soup = get_soup_from_url(url)
    if courses_list_soup:
        courses_urls = courses_list_soup.find_all('loc', limit=amount)
        return [tag.text for tag in courses_urls]


def parse_courses(urls_list):
    """
    :param urls_list: list of courses urls
    :return: list of dictionaries with courses info
    """
    courses_list = []
    for url in urls_list:
        course_soup = get_soup_from_url(url)
        courses_list.append(get_course_info(course_soup))
    return courses_list


def get_course_info(course_soup):
    """
    :param course_soup: soup of course main page
    :return: dictionary with course info. Keys: name, language, weeks,
     average_score, start_date
    """
    course_dictionary = {}

    try:
        course_dictionary['name'] = course_soup.find(
            class_=search_classes['course_name']).text
    except AttributeError:
        course_dictionary['name'] = '-'

    try:
        course_dictionary['language'] = course_soup.find(
            class_=search_classes['course_language']).text
    except AttributeError:
        course_dictionary['language'] = '-'

    try:
        course_dictionary['weeks'] = course_soup.find(
            class_=search_classes['weeks_tag']).text
    except AttributeError:
        course_dictionary['weeks'] = '-'

    try:
        weeks_list = course_soup.find(
            class_=search_classes['weeks_tag']).find_all(
            class_=search_classes['week_tag'])
        course_dictionary['weeks'] = len(weeks_list)
    except AttributeError:
        course_dictionary['weeks'] = '-'

    try:
        script_tag = course_soup.find('script',
                                      text=re.compile('window.App')).text
        course_dictionary['average_score'] = re.findall(
            r'"averageFiveStarRating":([\d.]+)', script_tag)[0]
    except (AttributeError, IndexError):
        course_dictionary['average_score'] = '-'

    try:
        course_script_vars = json.loads(course_soup.find(
            'script', type='application/ld+json').text.strip())
        course_dictionary['start_date'] = course_script_vars[
            'hasCourseInstance'][0]['startDate']
    except (AttributeError, IndexError, KeyError):
        course_dictionary['start_date'] = '-'
    return course_dictionary


def output_courses_info_to_xlsx(workbook, file_path, courses_info,
                                name_column_width, language_column_width):
    """
    :param workbook: openpyxl Workbook object
    :param file_path: file path for file saving
    :param courses_info: list of dictionaries with courses info
    :param name_column_width:
    :param language_column_width:
    :return: None
    """
    sheet = workbook.active
    for row, course in enumerate(courses_info):
        sheet.cell(row=row + 2, column=1).value = textwrap.fill(
            course['name'], name_column_width)
        sheet.cell(row=row + 2, column=2).value = textwrap.fill(
            course['language'], language_column_width)
        sheet.cell(row=row + 2, column=3).value = course['start_date']
        sheet.cell(row=row + 2, column=4).value = course['weeks']
        sheet.cell(row=row + 2, column=5).value = course['average_score']
    file_path = os.path.join(file_path)
    workbook.save(file_path)


def setup_excel_workbook(name_width, language_width, date_width=16,
                         duration_width=9, average_score_width=16):
    """
    Prepares excel workbook for filling
    :param name_width:
    :param language_width:
    :param date_width:
    :param duration_width:
    :param average_score_width:
    :return: openpyxl's Workbook object
    """
    workbook = Workbook()
    sheet = workbook.active
    sheet.column_dimensions['A'].width = name_width
    sheet.column_dimensions['B'].width = language_width
    sheet.column_dimensions['C'].width = date_width
    sheet.column_dimensions['D'].width = duration_width
    sheet.column_dimensions['E'].width = average_score_width
    sheet.cell(row=1, column=1).value = 'Name of course'
    sheet.cell(row=1, column=2).value = 'Language'
    sheet.cell(row=1, column=3).value = 'Date of start'
    sheet.cell(row=1, column=4).value = 'Duration'
    sheet.cell(row=1, column=5).value = 'Average score'
    return workbook


def parse_arguments():
    parser = argparse.ArgumentParser(description='Write Coursera courses info'
                                                 ' to xlsx')
    parser.add_argument('amount', nargs='?', default=20,
                        type=int, help='amount of courses to parse')
    parser.add_argument('output', nargs='?', default='courses.xlsx', type=str,
                        help='name of output file')
    parser.add_argument('name_width', nargs='?', default=40, type=int,
                        help='width of name column in spreadsheet')
    parser.add_argument('lang_width', nargs='?', default=40, type=int,
                        help='width of language column in spreadsheet')
    arguments = parser.parse_args()
    return arguments.amount, arguments.name_width, arguments.lang_width,\
        arguments.output


if __name__ == '__main__':
    courses_amount, name_col_width, language_col_width, output_file_name =\
        parse_arguments()

    if '.xlsx' not in output_file_name:
        output_file_name += '.xlsx'

    courses_urls_list = get_courses_list(coursera_xml_feed_url, courses_amount)
    print('---Getting course list...')

    if courses_urls_list:
        courses = parse_courses(courses_urls_list)
        print('---Parsing courses info. This can take a while (few seconds for'
              ' every course).')

        wb = setup_excel_workbook(name_col_width, language_col_width)
        output_courses_info_to_xlsx(wb, output_file_name, courses,
                                    name_col_width, language_col_width)
        print('---Success: {} is ready!'.format(output_file_name))
    else:
        print('!!! Something went wrong: check your internet connection;'
              ' also coursera xml feed may changed')
