"""Расчёт и текстовая интерпретация износа оборудования."""
from datetime import date


def wear(item):
    """Оценочный износ в процентах по сроку службы, а не результат экспертизы."""
    try:
        bought = date.fromisoformat(str(item["bought"]))
        life_years = float(item["life_years"])
        if life_years <= 0:
            return 0
        age = max(0, (date.today() - bought).days / 365.2425)
        return round(min(100, age / life_years * 100))
    except (TypeError, ValueError, KeyError, ZeroDivisionError):
        return 0


def wear_level(percent):
    """Короткая маркировка степени износа для интерфейса."""
    percent = int(percent or 0)
    if percent >= 80:
        return "Критический"
    if percent >= 60:
        return "Высокий"
    if percent >= 30:
        return "Умеренный"
    return "Низкий"


def wear_description(percent, bought=None, life_years=None):
    """Пояснение пользователю; оценка основана только на возрасте и сроке службы."""
    percent = int(percent or 0)
    if percent >= 80:
        summary = "Ресурс близок к исчерпанию. Рекомендуется техническая диагностика и решение о ремонте или замене."
    elif percent >= 60:
        summary = "Значительная часть расчётного ресурса использована. Рекомендуется усилить контроль состояния."
    elif percent >= 30:
        summary = "Оборудование находится в средней части расчётного срока службы. Продолжайте плановые проверки."
    else:
        summary = "Расчётный износ невысокий. Продолжайте обычное техническое обслуживание."
    details = []
    if bought:
        details.append(f"Дата приобретения: {bought}.")
    try:
        details.append(f"Расчётный срок службы: {float(life_years):g} лет.")
    except (TypeError, ValueError):
        pass
    return " ".join(details + [summary])
