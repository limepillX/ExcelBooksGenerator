import random
from openpyxl import Workbook

print('Программа начала работу, не открывайте создаваемый (или перемоздаваемый) файл.')
print('Если он открыт - закройте без сохранения!\n')

# Константы (их менять можно без проблем)
AMOUNT_OF_ROWS = 30_000  # Количество строчек, которые создаст файл

START = ['The book', 'It happened', 'Wonderful story',
         'Interesting situation happened', 'The life', 'Story', 'Ballad', 'Fairy tale', 'Secrets']  # Первые фразы (будут случайно комбинироваться с продолжениями)

START_2 = ['about', 'regarding to', 'conserning about', 'around',
           'throughout', 'all over about']  # Продолжения первых фраз

PLOT_SCHEMES = ['Фантастика', 'Детектив', 'Любовный роман', 'Триллер', 'Приключения',
                'Боевик', 'Исторический роман', 'Мистика', 'Фентези', 'Ужасы', 'Научная фантастика']  # Сюжетные схемы

QUALITY = ['Трагический', 'Сатирический', 'Комический',
           'Эпический', 'Идиллический']  # Ведущее эстетическое качество

Y_N = ['Да', 'Нет', 'Неизвестно']  # Основан ли на реальных событиях

# Открываем excel файл
workbook = Workbook()
sheet = workbook.active

# Вписываем изначальные значения
sheet["A1"] = "Имя"
sheet["B1"] = "Код товара"
sheet["C1"] = "Цена"
sheet["D1"] = "Объем"
sheet["E1"] = "Год выпуска"
sheet["F1"] = "Сюжетные схемы"
sheet["G1"] = "Ведущее эстетическое качество"
sheet["H1"] = "На реальных собыытиях"

# Открываем файлы со словами
adjectives = open('adj.txt', 'r')
nouns = open('nouns.txt', 'r')


# Генератор, при каждом вызове возвращает следующий код товара
def code_gen():
    i = 1
    while True:
        # Возвращаем строку из 8 цифр, недостающее количество заполняем нулями
        yield f'{str(i).rjust(8, "0")}'
        i += 1
gen = code_gen()


# Цикл, создаёт строчку в таблице
for i in range(AMOUNT_OF_ROWS):
    # Считыаем по одной строчки из файла с прилагательными
    current_adj = adjectives.readline().lower().strip()
    # Считыаем по одной строчки из файла с существительными
    current_noun = nouns.readline().lower().strip()

    # Проверяем не кончились ли слова в файах. Если кончились, то открываем файлы сначала
    if not current_adj:
        adjectives.close()
        adjectives = open('adj.txt', 'r')

    if not current_noun:
        nouns.close()
        nouns = open('nouns.txt', 'r')

    # Случайно выбираем значения, заносим их в строчку
    # random.choice берёт случайный элемент из массива
    sheet[f'A{i+2}'] = f'{random.choice(START)} {random.choice(START_2)} the {current_adj} {current_noun}'
    sheet[f'B{i+2}'] = gen.__next__() # получаем следующий номер товара (генераторы в Python работают через __next__)
    # Случайное число от 10 до 30
    sheet[f'C{i+2}'] = f'{random.randint(10, 30)}$'
    sheet[f'D{i+2}'] = f'{random.randint(25, 2_000)}'
    sheet[f'E{i+2}'] = f'{random.randint(1920, 2023)}'
    sheet[f'F{i+2}'] = f'{random.choice(PLOT_SCHEMES)}'
    sheet[f'G{i+2}'] = f'{random.choice(QUALITY)}'
    sheet[f'H{i+2}'] = f'{random.choice(Y_N)}'

# Завершаем работу программы
workbook.save(filename="table1.xlsx")  # Сохраняем excel файл

# Закрываем файлы
adjectives.close()
nouns.close()

print('Работа программы завершена, можно открывать файл.')
