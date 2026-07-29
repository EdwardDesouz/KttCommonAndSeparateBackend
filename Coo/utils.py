import re
def to_float(val):
    try:
        if val is None:
            return 0.0
        val = str(val).strip()
        if val in ['', 'N/A', '--', 'null']:
            return 0.0
        val = re.sub(r',', '', val)
        return float(val)
    except:
        return 0.0

def get_val(mapped, key, default=""):
    val = mapped.get(key, "")
    if val == "" or val is None:
        return default
    return val