from CustomOlympicParser import CustomOlympicParser

def main():
    url = "https://olimpiada.ru/activities"

    parser = CustomOlympicParser()
    markup = parser.getRequestPage(url=url)

    if(markup == None):
        return 0

    nodeElementList = parser.getNodeElementList(markup)
    preparedData = parser.parseDataFromNodeList(nodeList=nodeElementList)

    response = parser.saveDataToExcel("OlimpiadaParse.xlsx", "Список олимпиад", preparedData)

    if response: 
        print("[OK] Данные успешно получены и сохранены")
    else:
        print("[ERROR] Не удалось получить и сохранить данные")

if __name__ == "__main__":
    main()