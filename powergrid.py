__author__ = 'Frederick'

import os

# Python with COM requires the pyWin32 extensions
import win32com.client
from win32com.client import VARIANT

import pythoncom

current_dir = os.path.dirname(os.path.abspath(__file__))
object = win32com.client.Dispatch('pwrworld.SimulatorAuto')

# The following function will determine if any errors are returned and print an appropriate message.
def check_result_for_error(sim_auto_output, message):
    if sim_auto_output[0] != '':
        print('Error: ' + sim_auto_output[0])
    else:
        print message

filename = current_dir + '\\B7FLAT.pwb'
check_result_for_error(object.OpenCase(filename), 'Case Open')
object_type = 'GEN'

# VARIANT is needed if passing in array of arrays. BOTH the field list
# and the value list must use this syntax. If not passing in arrays of arrays, the
# standard list format can be used. Passing out arrays of arrays from SimAuto in the
# output parameter seems to work OK with Python.

field_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, ['BushNum', 'GenID', 'GenMW', 'GenAGCAble'])
all_value_array = [None] * 2
all_value_array[0] = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, [1, "1", 300, "NO"])
all_value_array[1] = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, [2, "1", 1400, "NO"])

check_result_for_error(object.ChangeParametersMultipleElement('GEN', field_array, all_value_array), 'Do Change')

filename = current_dir + '\\B7FLAT_changed.pwb'
check_result_for_error(object.SaveCase(filename, 'PWB', True), 'Save case')

del object
object = None
