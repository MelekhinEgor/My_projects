import win32com.client


def get_ncad_block_counts(ncad_layout):
    """
    Функция анализирующая число вхождения блоков.
    Возвращает словарь, ключомами являются наименования блоков, а значениями - числа вхождения блоков.

    :param ncad_layout: layout на котором необходимо произвести анализ числа вхождения блоков
    :return: словарь блоков, осортированный по наименованиям блоков
    """
    ncad_block_counts = {}
    ncad_blocks = ncad_layout.Block
    for one_block in ncad_blocks:
        if one_block.ObjectName == "AcDbBlockReference":
            one_block_name = one_block.Name
            if one_block_name in ncad_block_counts:
                ncad_block_counts[one_block_name] += 1
            else:
                ncad_block_counts[one_block_name] = 1
    sorted_ncad_block_counts = sorted(ncad_block_counts.items(), key=lambda x: x[0], reverse=False)
    return sorted_ncad_block_counts


def get_ncad_sorted_line_lengths(ncad_layout):
    """
    Функция анализирующая полилинии.
    Происходит сортировка полилиний по их принадлежности к слоям и подсчет суммарной длины
    Результатом является словарь полилиний, где ключами являются названия слоев, а значениями - сумманые длины линий
    в рамках одного слоя.

    :param ncad_layout: layout на котором необходимо произвести анализ полилиний
    :return: словарь полилиний, отсортированный по уменьшению суммы длин
    """
    ncad_line_lengths = {}
    ncad_blocks = ncad_layout.Block
    for one_line in ncad_blocks:
        if one_line.ObjectName == "AcDbPolyline":
            ncad_line_layer_name = one_line.Layer
            if ncad_line_layer_name in ncad_line_lengths:
                ncad_line_lengths[ncad_line_layer_name] += one_line.Length
            else:
                ncad_line_lengths[ncad_line_layer_name] = one_line.Length
    sorted_ncad_line_lengths = sorted(ncad_line_lengths.items(), key=lambda x: x[1], reverse=True)
    return sorted_ncad_line_lengths


def get_ncad_sorted_text_lengths(ncad_layout):
    """
    Функция анализирующая однострочный текст.
    Происходит сортировка однострочного текста по принадлежности к слою и подсчет количества символов на слое.
    Результатом явялется словарь однострочных текстов, где ключами являются названия слоев, а значениями - суммарное
    количество текстовых символов в рамках одного слоя.

    :param ncad_layout: layout на котором необходимо произвести анализ полилиний
    :return: слвоарь однострочных текстов, отсортированный по уменьшению количества текстовых символов
    """
    ncad_text_lengths = {}
    ncad_blocks = ncad_layout.Block
    for one_text in ncad_blocks:
        if one_text.ObjectName == "AcDbText":
            ncad_text_layer_name = one_text.Layer
            if ncad_text_layer_name in ncad_text_lengths:
                ncad_text_lengths[ncad_text_layer_name] += len(one_text.TextString)
            else:
                ncad_text_lengths[ncad_text_layer_name] = len(one_text.TextString)
    sorted_ncad_text_lengths = sorted(ncad_text_lengths.items(), key=lambda x: x[1], reverse=True)
    return sorted_ncad_text_lengths


def get_ncad_sorted_hatch_areas(ncad_layout):
    """
    Функция анализирующая штриховки.
    Происходит сортировка штриховок по принадлежности к слою и подсчет площадей штриховок на слое.
    Результатом является словарь штриховок, где ключами являются названия слоев, а значениями - суммарные площади всех
    штриховок в рамках одного слоя.

    :param ncad_layout: layout на котором необходимо произвести анализ полилиний
    :return: словарь штриховок, отсортированный по уменьшению суммарной площади
    """
    ncad_hatch_areas = {}
    ncad_blocks = ncad_layout.Block
    for one_hatch in ncad_blocks:
        if one_hatch.ObjectName == "AcDbHatch":
            ncad_hatch_layer_name = one_hatch.Layer
            if ncad_hatch_layer_name in ncad_hatch_areas:
                ncad_hatch_areas[ncad_hatch_layer_name] += one_hatch.Area
            else:
                ncad_hatch_areas[ncad_hatch_layer_name] = one_hatch.Area
    sorted_ncad_hatch_areas = sorted(ncad_hatch_areas.items(), key=lambda x: x[1], reverse=True)
    return sorted_ncad_hatch_areas


def create_table(data, insert_point, ncad_Layout, col_1_name="Колонка 1", col_2_name="Колонка 2"):
    """
    Функция формирующая блок таблицы на выбранном layout в nanoCAD.
    Формируемые таблицы состоят из двух колонок с шириной 50 мм каждая.
    У полученной таблицы удаляется строка названия таблицы, чтобы это отключить необходимо закомментировать 114 строку,
    а в 109 строке в кавычках указать название таблицы, которое будет выводиться на лист.

    :param data: данные для заполнения таблицы
    :param insert_point: верхний правый угол таблицы, являющийся точкой ее вставки
    :param ncad_Layout: layout на котором необходимо разместить блок таблицы
    :param col_1_name: наименование первой колонки, по умолчанию стоит - Колонка 1
    :param col_2_name: наименование второй колонки, по умолчанию стоит - Колонка 2
    :return: графическое представление в nanoCAD на выбранном layout
    """
    table = ncad_Layout.Block.AddTable(insert_point, len(data) + 2, 2, 8, 50)
    table.SetText(0, 0, "")
    table.SetText(1, 0, col_1_name)
    table.SetText(1, 1, col_2_name)
    for i, row in enumerate(data, start=2):
        table.SetText(i, 0, row[0])
        table.SetText(i, 1, str(row[1]))
    table.DeleteRows(0, 1)


def create_text(insertion_text_point, text_width, text):
    """
    Функция формирующая блок MText на выбранном layout в nanoCAD

    :param insertion_text_point: точка верхнего правого угла, точка вставки текста
    :param text_width: ширина создаваемого текстового поля
    :param text: текст, который будет отображаться на выбранном layout
    :return: графическое представление в nanoCAD на выбранном layout
    """
    ncad_Layout.Block.AddMText(insertion_text_point, text_width, text)


"""
------------------------------------------------------------------------------------------------------------------------
Далее представлен основной код, который подключается к активному окну nanoCAD
и выполняет заложенные в него функции, а именно:

Из пространства модели (ModelSpace) получает количественные характеристики:

- количество Вхождений блока каждого типа;
- суммарная длина всех линий с сортировкой по слоям;
- суммарное количество текстовых символов во всех Однострочных текстах с сортировкой по слоям;
- суммарная площадь всей штриховки с сортировкой по слоям

Полученные количественные характеристики формляет в Таблицы и размещает в пространстве листа "Для вставки таблиц" ,
также делает к каждой таблице элементом MText заголовок (равный заданию из выделенных подпунктов выше)
------------------------------------------------------------------------------------------------------------------------
"""
ncad_app = win32com.client.Dispatch("nanoCAD.Application")
if ncad_app is not None:
    ncad_doc = ncad_app.ActiveDocument
    if ncad_doc is not None:
        for one_layout_index in range(0, ncad_doc.Layouts.Count, 1):
            ncad_Layout = ncad_doc.Layouts.Item(one_layout_index)
            if ncad_Layout.Name == "Model":
                ncad_block_counts = get_ncad_block_counts(ncad_Layout)
                ncad_line_lengths = get_ncad_sorted_line_lengths(ncad_Layout)
                ncad_text_lengths = get_ncad_sorted_text_lengths(ncad_Layout)
                ncad_hatch_areas = get_ncad_sorted_hatch_areas(ncad_Layout)

            if ncad_Layout.Name == "Для вставки таблтиц":
                size_of_layout = ncad_Layout.GetPaperSize()
                # Создание и размещения текста и таблицы с анализом блоков
                table_insert_point = str(size_of_layout[0] / 12) + "," + str(size_of_layout[1] - 80) + ",0"
                create_table(ncad_block_counts, table_insert_point, ncad_Layout, "Тип блока", "Количество")
                text_insert_point = str(size_of_layout[0] / 12) + "," + str(size_of_layout[1] - 60) + ",0"
                create_text(text_insert_point, 100, "Количество вхождений блока каждого типа")
                # Создание и размещения текста и таблицы с анализом полилиний
                table_insert_point = str(size_of_layout[0] / 12 + 150) + "," + str(size_of_layout[1] - 80) + ",0"
                create_table(ncad_line_lengths, table_insert_point, ncad_Layout, "Cлой", "Длина")
                text_insert_point = str(size_of_layout[0] / 12 + 150) + "," + str(size_of_layout[1] - 60) + ",0"
                create_text(text_insert_point, 100, "Суммарная длина всех линий с сортировкой по слоям")
                # Создание и размещения текста и таблицы с анализом однострочного текста
                table_insert_point = str(size_of_layout[0] / 12 + 300) + "," + str(size_of_layout[1] - 80) + ",0"
                create_table(ncad_text_lengths, table_insert_point, ncad_Layout, "Слой", "Количество")
                text_insert_point = str(size_of_layout[0] / 12 + 300) + "," + str(size_of_layout[1] - 45) + ",0"
                create_text(text_insert_point, 100, "Суммарное количество текстовых символов во всех Однострочных "
                                                    "текстах с сортировкой по слоям")
                # Создание и размещения текста и таблицы с анализом штриховок
                table_insert_point = str(size_of_layout[0] / 12 + 450) + "," + str(size_of_layout[1] - 80) + ",0"
                create_table(ncad_hatch_areas, table_insert_point, ncad_Layout, "Слой", "Площадь")
                text_insert_point = str(size_of_layout[0] / 12 + 450) + "," + str(size_of_layout[1] - 60) + ",0"
                create_text(text_insert_point, 100, "Суммарная площадь всей штриховки с сортировкой по слоям")
    else:
        print("Doc not runing")
else:
    print("App not runing")
