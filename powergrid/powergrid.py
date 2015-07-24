# -*- coding: utf-8 -*-

import os
import numpy as np

# Python with COM requires the pyWin32 extensions
import win32com.client
from win32com.client import VARIANT

import pythoncom

current_dir = os.path.dirname(os.path.abspath(__file__))
# Create PowerWorld COM object
par_sim_auto = win32com.client.Dispatch('pwrworld.SimulatorAuto')

# The following function will determine if any errors are returned and print an appropriate message.
def check_result_for_error(sim_auto_output, message):
    if sim_auto_output[0] != '':
        print('Error: ' + sim_auto_output[0])
    else:
        print message
        return sim_auto_output

# Load sample case
filename = current_dir + '\\resources\\sampleCase.pwb'

# initializePWCase
check_result_for_error(par_sim_auto.OpenCase(filename), 'Case Open')
check_result_for_error(par_sim_auto.RunScriptCommand('EnterMode(RUN)'), 'Enter Mode RUN')

# Save state from before we switch
check_result_for_error(par_sim_auto.SaveState(), 'Save State')

# Run OPF
check_result_for_error(par_sim_auto.RunScriptCommand('SolvePrimalLP'), 'Solve Primal LP')

# getBranchState
change_status_field_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, ['busnum', 'busnum:1', 'LineCircuit', 'LineStatus'])
output_lines = check_result_for_error(par_sim_auto.GetParametersMultipleElement('Branch', change_status_field_array, ' '), 'Branch')
output_lines = np.array(output_lines[1]).T
output_flattened = output_lines.flatten()

del par_sim_auto
par_sim_auto = None
