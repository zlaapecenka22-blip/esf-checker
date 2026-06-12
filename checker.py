import pandas as pd
from columns import COLUMNS


def compare_files(file_1c, file_esf):

    df_1c = pd.read_excel(file_1c)
    df_esf = pd.read_excel(file_esf)

    results = []

    # -------------------------
    # Дубликаты в 1С
    # -------------------------

    duplicates_1c = df_1c[
        df_1c.duplicated(
            subset=[COLUMNS["number"]],
            keep=False
        )
    ]

    for _, row in duplicates_1c.iterrows():

        results.append({
            "НомерСФ": row[COLUMNS["number"]],
            "Статус": "Ошибка",
            "Тип": "Дубликат",
            "Комментарий": "Дубликат в 1С"
        })

    # -------------------------
    # Дубликаты в ЭСФ
    # -------------------------

    duplicates_esf = df_esf[
        df_esf.duplicated(
            subset=[COLUMNS["number"]],
            keep=False
        )
    ]

    for _, row in duplicates_esf.iterrows():

        results.append({
            "НомерСФ": row[COLUMNS["number"]],
            "Статус": "Ошибка",
            "Тип": "Дубликат",
            "Комментарий": "Дубликат в ЭСФ"
        })

    # -------------------------
    # Сверка документов
    # -------------------------

    for _, row in df_1c.iterrows():

        number = row[COLUMNS["number"]]

        if number in duplicates_1c[COLUMNS["number"]].values:
            continue

        match = df_esf[
            df_esf[COLUMNS["number"]] == number
        ]

        if len(match) == 0:

            results.append({
                "НомерСФ": number,
                "Статус": "Ошибка",
                "Тип": "Нет в ЭСФ",
                "Комментарий": "Документ отсутствует в ЭСФ"
            })

            continue

        esf_row = match.iloc[0]

        errors = []

        # Дата

        date_1c = pd.to_datetime(
            row[COLUMNS["date"]]
        ).date()

        date_esf = pd.to_datetime(
            esf_row[COLUMNS["date"]]
        ).date()

        if date_1c != date_esf:

            errors.append(
                f"Дата: 1С={date_1c}, ЭСФ={date_esf}"
            )

        # Сумма

        if float(row[COLUMNS["amount"]]) != float(
            esf_row[COLUMNS["amount"]]
        ):

            errors.append(
                f"Сумма: 1С={row[COLUMNS['amount']]}, ЭСФ={esf_row[COLUMNS['amount']]}"
            )

        # НДС

        if float(row[COLUMNS["vat"]]) != float(
            esf_row[COLUMNS["vat"]]
        ):

            errors.append(
                f"НДС: 1С={row[COLUMNS['vat']]}, ЭСФ={esf_row[COLUMNS['vat']]}"
            )

        if errors:

            results.append({
                "НомерСФ": number,
                "Статус": "Ошибка",
                "Тип": "Расхождение",
                "Комментарий": "; ".join(errors)
            })

        else:

            results.append({
                "НомерСФ": number,
                "Статус": "OK",
                "Тип": "Совпадение",
                "Комментарий": "Полностью совпадает"
            })

    # -------------------------
    # Нет в 1С
    # -------------------------

    for _, row in df_esf.iterrows():

        number = row[COLUMNS["number"]]

        if number in duplicates_esf[COLUMNS["number"]].values:
            continue

        match = df_1c[
            df_1c[COLUMNS["number"]] == number
        ]

        if len(match) == 0:

            results.append({
                "НомерСФ": number,
                "Статус": "Ошибка",
                "Тип": "Нет в 1С",
                "Комментарий": "Документ отсутствует в 1С"
            })

    return results