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
        course = Course(course_soup)
        courses_list.append(course.get_course_info())
    return courses_list


class Course:
    # I had to create this class because my course parse function had MacCabe
    # function complexity score of 13 and i didn't want to pass heavy soup
    # object to every single parse function
    def __init__(self, soup):
        self.soup = soup

    def get_name(self):
        try:
            return self.soup.find(
                class_=search_classes['course_name']).text
        except AttributeError:
            return '-'

    def get_language(self):
        try:
            return self.soup.find(
                class_=search_classes['course_language']).text
        except AttributeError:
            return '-'

    def get_duration(self):
        try:
            weeks_list = self.soup.find(
                class_=search_classes['weeks_tag']).find_all(
                class_=search_classes['week_tag'])
            return len(weeks_list)
        except AttributeError:
            return '-'

    def get_average_score(self):
        try:
            script_tag = self.soup.find('script',
                                        text=re.compile('window.App')).text
            return re.findall(r'"averageFiveStarRating":([\d.]+)',
                              script_tag)[0]
        except (AttributeError, IndexError):
            return '-'

    def get_start_date(self):
        try:
            course_script_vars = json.loads(self.soup.find('script',
                                            type='application/ld+json').text.
                                            strip())
            return course_script_vars['hasCourseInstance'][0]['startDate']
        except (AttributeError, IndexError, KeyError):
            return '-'

    def get_course_info(self):
        """
        :return: dictionary with course info. Keys: name, language, weeks,
         average_score, start_date
        """
        course_dictionary = {'name': self.get_name(),
                             'language': self.get_language(),
                             'weeks': self.get_duration(),
                             'average_score': self.get_average_score(),
                             'start_date': self.get_start_date()}
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
    """
    :return: tuple of length 4: courses amount, name column width, language
      column width, output file name
    """
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
