# -*- coding: utf-8 -*-

import os
import numpy as np
import multiprocessing as mp
import timeit

# Python with COM requires the pyWin32 extensions
import win32com.client
from win32com.client import VARIANT

import pythoncom

current_dir = os.path.dirname(os.path.abspath(__file__))
filename = current_dir + '\\resources\\sampleCase.pwb'

# The following function will determine if any errors are returned and print an appropriate message.
def check_result_for_error(sim_auto_output, message):
    if sim_auto_output[0] != '':
        print('Error: ' + sim_auto_output[0])
    else:
        print message
        return sim_auto_output

results = []
def log_result(result):
    results.append(result)

def serial():
    for i in range(2):
        log_result(load_sample_case())

def multiprocess(processes):
    # Create pool of x processes
    pool = mp.Pool(processes=processes)
    for i in range(2):
        # Spawn processes immediately without waiting for old one ot finish
        pool.apply_async(load_sample_case, args=(), callback=log_result)
    pool.close()
    pool.join()
    # print(results)


def load_sample_case():
    # Create PowerWorld COM object
    par_sim_auto = win32com.client.Dispatch('pwrworld.SimulatorAuto')

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

    par_sim_auto = None
    del par_sim_auto

    return output_flattened

benchmarks = []

if __name__ == '__main__':
    # benchmarks.append(timeit.Timer('serial()', 'from __main__ import serial').timeit(number=1))
    # benchmarks.append(timeit.Timer('multiprocess(1)', 'from __main__ import multiprocess').timeit(number=1))
    benchmarks.append(timeit.Timer('multiprocess(2)', 'from __main__ import multiprocess').timeit(number=1))
    # benchmarks.append(timeit.Timer('multiprocess(3)', 'from __main__ import multiprocess').timeit(number=1))
    # benchmarks.append(timeit.Timer('multiprocess(4)', 'from __main__ import multiprocess').timeit(number=1))
    print(benchmarks)
