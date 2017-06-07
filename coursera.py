from lxml import html
from bs4 import BeautifulSoup
from openpyxl import Workbook
import requests
import random


def get_courses_list(number_of_courses):
    url_link = 'https://www.coursera.org/sitemap~www~courses.xml'
    full_list_courses = []
    course_slug = []
    xml_feed = requests.get(url_link)
    xml_feed = xml_feed.content
    list_courses = html.fromstring(xml_feed)

    for element in list_courses.xpath('//url'):
        full_list_courses.append(element.getchildren()[0].text)

    course_slug = random.sample(full_list_courses, number_of_courses)
    return course_slug


def get_html_content(url_link):
    html_content = requests.get(url_link).content
    return html_content


def get_course_info(html_content):
    content = BeautifulSoup(html_content, 'html.parser')
    name_course = content('h1', class_='title display-3-text')[0].get_text()
    language_course = content('div', class_='rc-Language')[0].get_text()
    start_date = content('div', class_='startdate rc-StartDateString'
                                       ' caption-text')[0].get_text()
    duration_of_course = len(content('div', class_='week-heading body-2-text'))
    try:
        rating_course = content('div', class_='ratings-text'
                                              ' headline-2-text')[0].get_text()
    except IndexError:
        rating_course = 'no data'

    course_info_list = (name_course,
                        language_course,
                        start_date,
                        duration_of_course,
                        rating_course)
    return course_info_list


def get_head_info(name='Course name',
                  language='Language',
                  start_date='Start date of the course',
                  duration='Duration week(s)',
                  rating='Rating'):
    head_info_list = [name, language, start_date, duration, rating]
    return head_info_list


def output_courses_info_to_xlsx(filepath,
                                courses_info_list,
                                head_info_list):
    work_book = Workbook()
    work_sheet = work_book.create_sheet()
    work_sheet.append(head_info_list)
    for course in courses_info_list:
        work_sheet.append(course)
    work_book.save(filepath)


if __name__ == '__main__':
    filepath = 'courses_info.xlsx'
    number_of_courses = 3
    courses_info_list = []
    course_slug = get_courses_list(number_of_courses)

    for url_link in course_slug:
        html_content = get_html_content(url_link)
        course_info_list = get_course_info(html_content)
        courses_info_list.append(course_info_list)

    head_info_list = get_head_info()
    output_courses_info_to_xlsx(filepath, courses_info_list, head_info_list)