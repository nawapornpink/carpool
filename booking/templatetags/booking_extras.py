from django import template

register = template.Library()


@register.filter
def dict_get(d, key):
    """ใช้ใน template: bookings_by_day|dict_get:day"""
    if d is None:
        return []
    return d.get(key, [])


@register.filter(name="split")
def split(value, sep=","):
    if value is None:
        return []
    return str(value).split(sep)
