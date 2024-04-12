import requests
from typing import Union
from openpyxl import Workbook
from bs4 import Tag, ResultSet, BeautifulSoup as bs

class CustomOlympicParser:
    """Custom parser for parsing olympics"""

    def getRequestPage(self, url : str) -> Union[str, None]:
        """Getting HTML-markup from url. If URL bad - returned None"""
        page = requests.get(url)

        if page.status_code == 200:
            return page.text


    def getNodeElementList(self, html_markup: str) -> Union[ResultSet, None]:
        """Getting node list from soup"""

        if html_markup:
            soup = bs(markup=html_markup, features="html.parser")
            return soup.find_all("div", class_="fav_olimp olimpiada")

    def parseDataFromNode(self, node: Tag) -> dict:
        """Prepares data to save from node list"""

        title = node.find("span", class_="headline")
        raiting = node.find("span", class_="pl_rating")
        description = node.find("a", class_="none_a black olimp_desc")
        studentClasses = node.find("span", class_="classes_dop")
        subjectTags = node.find_all("span", class_="subject_tag")
        url = node.find("a", class_="none_a black").get_attribute_list("href")[0]
        
        olimpicData = {
            "title": title.text if title is not None else "",
            "raitng": raiting.text if raiting is not None else "",
            "description": description.text if description is not None else "",
            "studentClasses": studentClasses.text if studentClasses is not None else "",
            "subjectTags": ", ".join(map(lambda x: x.text.strip(), subjectTags)) if subjectTags is not None else "",
            "url": f"https://olimpiada.ru{url}" if url is not None else ""
        }

        return olimpicData

def main():
    url = "https://olimpiada.ru/activities"

    parser = CustomOlympicParser()
    markup = parser.getRequestPage(url=url)

    if(markup == None):
        return 0

    nodeElementList = parser.getNodeElementList(markup)

    for node in nodeElementList:
        data = parser.parseDataFromNode(node)
        print(data)

if __name__ == "__main__":
    main()