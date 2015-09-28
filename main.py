#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import numpy as np
import timeit
import cProfile

from powergrid.powergrid import PowerGrid
# Python with COM requires the pyWin32 extensions
from win32com.client import VARIANT

import pythoncom

current_dir = os.path.dirname(os.path.abspath(__file__))
filename = current_dir + '\\powergrid\\resources\\sampleCase.pwb'
pw = PowerGrid(15)

# The following function will determine if any errors are returned and print an appropriate message.
def check_result_for_error(sim_auto_output, message):
    if sim_auto_output[0] != '':
        print('Error: ' + sim_auto_output[0])
    else:
        #print(message)
        return sim_auto_output

def create_pw_pool():
    # Create 8 PowerWorld COM objects
    pw.create_pw_pool()

@pw.threaded
def threaded_func(parameter, com_id=None, auto_sim=None):
    #print('%s: Starting new thread: %s' % (com_id, parameter))
    # initializePWCase
    check_result_for_error(auto_sim.OpenCase(filename), 'Case Open')
    check_result_for_error(auto_sim.RunScriptCommand('EnterMode(RUN)'), 'Enter Mode RUN')

    # Save state from before we switch
    check_result_for_error(auto_sim.SaveState(), 'Save State')

    # Run OPF
    check_result_for_error(auto_sim.RunScriptCommand('SolvePrimalLP'), 'Solve Primal LP')

    # getBranchState
    change_status_field_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, ['busnum', 'busnum:1', 'LineCircuit', 'LineStatus'])
    output_lines = check_result_for_error(auto_sim.GetParametersMultipleElement('Branch', change_status_field_array, ' '), 'Branch')
    output_lines = np.array(output_lines[1]).T
    output_flattened = output_lines.flatten()

    # Store the results along with the COM id
    return com_id, output_flattened

@pw.threaded # replica of threaded_func, except to remove print statement and OpenCase SimAuto function
def testPerformanceFunc(parameter, com_id=None, auto_sim=None):
    check_result_for_error(auto_sim.RunScriptCommand('EnterMode(RUN)'), 'Enter Mode RUN')

    # Save state from before we switch
    check_result_for_error(auto_sim.SaveState(), 'Save State')

    # Run OPF
    check_result_for_error(auto_sim.RunScriptCommand('SolvePrimalLP'), 'Solve Primal LP')

    # getBranchState
    change_status_field_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, ['busnum', 'busnum:1', 'LineCircuit', 'LineStatus'])
    output_lines = check_result_for_error(auto_sim.GetParametersMultipleElement('Branch', change_status_field_array, ' '), 'Branch')
    output_lines = np.array(output_lines[1]).T
    output_flattened = output_lines.flatten()

    # Store the results along with the COM id
    return com_id, output_flattened

def multiprocess():
    results = threaded_func("foo") # run it one time to load case
    for i in range(1,1000):
        results = testPerformanceFunc("foo")
        while not results.empty():
            result = results.get()
            #print(result)
    
    pw.kill_com_objects()


benchmarks = []
if __name__ == '__main__':
    benchmarks.append(timeit.Timer('create_pw_pool()', 'from __main__ import create_pw_pool').timeit(number=1))
    #benchmarks.append(timeit.Timer('multiprocess()', 'from __main__ import multiprocess').timeit(number=1))
    cProfile.run('multiprocess()')
    print(benchmarks)
