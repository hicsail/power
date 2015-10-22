#!/usr/bin/python
# -*- coding: utf-8 -*-

import os
import numpy as np
import timeit

from power.power import Power
# Python with COM requires the pyWin32 extensions
from win32com.client import VARIANT

import pythoncom

current_dir = os.path.dirname(os.path.abspath(__file__))
filename = current_dir + '\\power\\resources\\sampleCase.pwb'
pw = Power(4)


# The following function will determine if any errors are returned and print an appropriate message.
def check_result_for_error(sim_auto_output, message):
    if sim_auto_output[0] != '':
        print('Error: ' + sim_auto_output[0])
    else:
        print(message)
        return sim_auto_output


def create_pw_pool():
    # Create 8 PowerWorld COM objects
    pw.create_pw_pool()


def threaded_func(thread_id=None, auto_sim=None):
    print('Starting new thread: %s' % thread_id)
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
    return output_flattened


def callback(task, result):
    print("**** Result from request #%s: %r" % (task.pw_id, result))


def multiprocess():
    results = pw.add_task(threaded_func, '0-3')
    print(results)
    pw.dismiss_threads('0-3')
    pw.reset()


benchmarks = []
if __name__ == '__main__':
    benchmarks.append(timeit.Timer('create_pw_pool()', 'from __main__ import create_pw_pool').timeit(number=1))
    benchmarks.append(timeit.Timer('multiprocess()', 'from __main__ import multiprocess').timeit(number=1))
    print(benchmarks)
