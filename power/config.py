import configparser
import os
import ast


def _build_dict():
    """
    Turn config object into dictionary.
    Config values are usually represented as strings, no matter the actual data type. You need to use config.getfloat()
    and others when appropriate to get the right values. This parses everything automatically for you and puts it in a
    flat dictionary (sections are not kept).

    :return: Parsed flat config file as dictionary
    """
    _obj = {}
    for section in conf.sections():
        options = conf.options(section)
        for option in options:
            _obj[option] = _parse_string(conf.get(section, option))
    return _obj


def _parse_string(s):
    """
    Parse string to correct data type.

    :param s: String to parse
    :return: Parsed string to either int, float, boolean or return the original value.
    """
    # Parse to boolean
    if s.lower() in ['true', '1', 'on', 'yes']:
        return True
    elif s.lower() in ['false', '0', 'off', 'no']:
        return False

    # Parse to int/float
    try:
        return ast.literal_eval(s)
    except (ValueError, SyntaxError):
        # Not int, float or boolean, return the original value
        return s

# Read config file
conf = configparser.ConfigParser()
# Get parent of parent dir, then append config.ini
conf.read(os.path.abspath(os.sep.join(__file__.split(os.sep)[:-2]) + os.sep + 'config.ini'))
_dict = _build_dict()


def get(key, fallback=None):
    """
    Retrieve value from config file, wrapper for config dictionary get()

    :param key: Key to get value for, will be converted to lowercase
    :param fallback: Optional fallback value if key doesn't exist
    :return: Value from config dict or fallback
    """
    return _dict.get(key.lower(), fallback)


def put(key, value):
    """
    Set value to config file

    :param key: Key to set value under, will be made lowercase
    :param value: Value to set
    """
    _dict[key.lower()] = value


def data():
    """
    Get full config dictionary, used when sending the entire state to newly connected clients.

    :return: Config dictionary
    """
    return _dict
