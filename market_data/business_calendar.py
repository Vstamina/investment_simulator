from datetime import date, datetime, timedelta

import holidays


def normalize_date(value):
    """
    Converte date, datetime ou string ISO em date.
    """
    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    if isinstance(value, str):
        return datetime.fromisoformat(value).date()

    raise ValueError(f"Tipo de data não reconhecido: {type(value)}")


def get_brazil_national_holidays(start_date, end_date):
    """
    Retorna feriados nacionais brasileiros no intervalo informado.
    Usa a biblioteca holidays.
    """
    start_date = normalize_date(start_date)
    end_date = normalize_date(end_date)

    years = range(start_date.year, end_date.year + 1)

    brazil_holidays = holidays.Brazil(years=years)

    return {
        holiday_date
        for holiday_date in brazil_holidays.keys()
        if start_date <= holiday_date <= end_date
    }


def is_business_day(current_date, holiday_set=None):
    """
    Verifica se a data é dia útil:
    - segunda a sexta;
    - exclui feriados nacionais.
    """
    current_date = normalize_date(current_date)

    if holiday_set is None:
        holiday_set = set()

    is_weekday = current_date.weekday() < 5

    return is_weekday and current_date not in holiday_set


def count_business_days(start_date, end_date, include_start=False):
    """
    Conta dias úteis entre duas datas.

    Por padrão, não inclui a data inicial.
    Exemplo:
    aplicação em 01/01 e vencimento em 02/01 conta o rendimento de 1 dia.
    """
    start_date = normalize_date(start_date)
    end_date = normalize_date(end_date)

    if end_date <= start_date:
        return 0

    holiday_set = get_brazil_national_holidays(start_date, end_date)

    current_date = start_date if include_start else start_date + timedelta(days=1)

    business_days = 0

    while current_date <= end_date:
        if is_business_day(current_date, holiday_set):
            business_days += 1

        current_date += timedelta(days=1)

    return business_days


def iter_business_days(start_date, end_date, include_start=False):
    """
    Gera os dias úteis do intervalo.
    """
    start_date = normalize_date(start_date)
    end_date = normalize_date(end_date)

    if end_date <= start_date:
        return

    holiday_set = get_brazil_national_holidays(start_date, end_date)

    current_date = start_date if include_start else start_date + timedelta(days=1)

    while current_date <= end_date:
        if is_business_day(current_date, holiday_set):
            yield current_date

        current_date += timedelta(days=1)


def count_calendar_days(start_date, end_date):
    """
    Conta dias corridos entre duas datas.
    Usado para IR regressivo.
    """
    start_date = normalize_date(start_date)
    end_date = normalize_date(end_date)

    if end_date <= start_date:
        return 0

    return (end_date - start_date).days