# -*- coding: utf-8 -*-

import os
import numpy as np
import timeit

# Python with COM requires the pyWin32 extensions
import win32com.client
from win32com.client import VARIANT

import pythoncom
from Queue import Queue
from threading import Thread

current_dir = os.path.dirname(os.path.abspath(__file__))
filename = current_dir + '\\resources\\sampleCase.pwb'
pw_objects = Queue()
max_objects = 2

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


def create_pw_pool():
    for i in range(max_objects):
        # Create COM object
        pw = win32com.client.Dispatch('pwrworld.SimulatorAuto')
        # Create stream that will hold COM object
        pw_stream = pythoncom.CreateStreamOnHGlobal()
        # Convert COM object into stream to allow re-usage
        pythoncom.CoMarshalInterface(pw_stream,
                                     pythoncom.IID_IDispatch,
                                     pw._oleobj_,
                                     pythoncom.MSHCTX_LOCAL,
                                     pythoncom.MSHLFLAGS_TABLESTRONG)
        # No need for the COM reference anymore now that it's a stream
        pw = None
        # Store the stream in a queue
        pw_objects.put(pw_stream)

def multiprocess(threads):
    for i in range(threads):
        worker = Thread(target=load_sample_case, args=(i, pw_objects,))
        worker.start()
    worker.join()

    clean_pw_queue()
    print '*** Done'


def clean_pw_queue():
    while not pw_objects.empty():
        pw = pw_objects.get()
        pythoncom.CoReleaseMarshalData(pw)
        pw = None
        pw_objects.task_done()
    pw_objects.mutex.acquire()
    pw_objects.queue.clear()
    pw_objects.all_tasks_done.notify_all()
    pw_objects.unfinished_tasks = 0
    pw_objects.mutex.release()


def load_sample_case(i, q):
    print '%s: Starting new thread' % i
    # Enable COM object access in this thread, but not others
    pythoncom.CoInitialize()
    # Get stream reference from queue
    pw_stream = q.get()
    # Make sure we're at the start of the stream, reset the pointer
    pw_stream.Seek(0, 0)
    # Unmarshal the stream, going back to the original interface
    pw_interface = pythoncom.CoUnmarshalInterface(pw_stream, pythoncom.IID_IDispatch)
    # And finally return the COM object that was created earlier
    pw = win32com.client.Dispatch(pw_interface)

    # initializePWCase
    check_result_for_error(pw.OpenCase(filename), 'Case Open')
    check_result_for_error(pw.RunScriptCommand('EnterMode(RUN)'), 'Enter Mode RUN')

    # Save state from before we switch
    check_result_for_error(pw.SaveState(), 'Save State')

    # Run OPF
    check_result_for_error(pw.RunScriptCommand('SolvePrimalLP'), 'Solve Primal LP')

    # getBranchState
    change_status_field_array = VARIANT(pythoncom.VT_VARIANT | pythoncom.VT_ARRAY, ['busnum', 'busnum:1', 'LineCircuit', 'LineStatus'])
    output_lines = check_result_for_error(pw.GetParametersMultipleElement('Branch', change_status_field_array, ' '), 'Branch')
    output_lines = np.array(output_lines[1]).T
    output_flattened = output_lines.flatten()

    log_result(output_flattened)

    # Revert stream back to start position
    pw_stream.Seek(0, 0)
    # Return stream back to queue
    q.put(pw_stream)
    # Indicate all work on this queue object is done. Without this the queue task counter would go up every time a
    # stream is re-added to the queue
    q.task_done()
    # Clean up COM reference
    pw = None
    # Indicate that no more COM objects will be called in this thread
    pythoncom.CoUninitialize()


benchmarks = []

if __name__ == '__main__':
    benchmarks.append(timeit.Timer('create_pw_pool()', 'from __main__ import create_pw_pool').timeit(number=1))
    # benchmarks.append(timeit.Timer('serial()', 'from __main__ import serial').timeit(number=1))
    # benchmarks.append(timeit.Timer('multiprocess(1)', 'from __main__ import multiprocess').timeit(number=1))
    benchmarks.append(timeit.Timer('multiprocess(4)', 'from __main__ import multiprocess').timeit(number=1))
    # benchmarks.append(timeit.Timer('multiprocess(3)', 'from __main__ import multiprocess').timeit(number=1))
    # benchmarks.append(timeit.Timer('multiprocess(4)', 'from __main__ import multiprocess').timeit(number=1))
    print(benchmarks)
