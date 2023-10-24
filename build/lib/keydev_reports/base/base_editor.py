

class BaseEditor:
    def __init__(self, file_name: str):
        self.file_name = file_name

    @staticmethod
    def get_search_text(search_text):
        """
        Статический метод для формирования текста поиска (ключа) на основе переданного текста.
        :param search_text: Текст, на основе которого формируется ключ для поиска.
        :return: Строка, представляющая ключ для поиска.
        """
        return '${' + search_text + '}'

    def get_filepath(self):
        """
        Метод для получения рабочей книги (workbook) класса.
        :return: Рабочая книга (workbook).
        """
        return self.file_name
