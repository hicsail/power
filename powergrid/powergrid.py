import win32com.client
from win32com.client import VARIANT

import sys
import pythoncom
import queue
from queue import Queue
from threading import Thread
from typing import Sequence, TypeVar


class PowerThread(Thread):
    """
    A thread that runs indefinitely and uses a queue to obtain more tasks.
    """
    T = TypeVar('T', Queue)

    def __init__(self, tasks: Sequence[T], results: Queue, pw_objects: Queue, **kwargs):
        Thread.__init__(self, **kwargs)
        self.setDaemon(1)
        self._results = results
        self._pw_object = pw_objects
        self._pw_id, self._pw, self._pw_stream = self.marshal_com()
        self._tasks = tasks
        self.start()

    def marshal_com(self):
        # Enable COM object access in this thread, but not others
        pythoncom.CoInitialize()
        # Get tuple of id and stream
        pw_data = self._pw_objects.get()
        # Get the ID
        pw_id = pw_data[0]
        # Get stream reference from queue
        pw_stream = pw_data[1]
        # Make sure we're at the start of the stream, reset the pointer
        pw_stream.Seek(0, 0)
        # Unmarshal the stream, going back to the original interface
        pw_interface = pythoncom.CoUnmarshalInterface(pw_stream, pythoncom.IID_IDispatch)
        # And finally return the COM object that was created earlier
        pw = win32com.client.Dispatch(pw_interface)

        return pw, pw_id, pw_stream

    def unmarshal_com(self):
        # Revert stream back to start position
        self._pw_stream.Seek(0, 0)
        # Return stream back to queue
        self._pw_objects.put((self._pw_id, self._pw_stream))
        # Indicate all work on this queue object is done. Without this the queue task counter would go up every time
        # a stream is re-added to the queue
        self._pw_objects.task_done()
        # Clean up COM reference
        self._pw = None
        # Indicate that no more COM objects will be called in this thread
        pythoncom.CoUninitialize()

    def run(self):
        while True:
            try:
                # Get task with non-blocking Queue call
                task = self._tasks[self._pw_id].get(False)
            except queue.Empty:
                continue
            else:
                try:
                    # Call task function and store results
                    result = task.f(*task.args, com_id=self._pw_id, auto_sim=self._pw, **task.kwargs)
                    self._results.put((task, result))
                except:
                    # Or store exception message if something went wrong
                    self._results.put((task, sys.exc_info()))


class PowerTask:
    def __init__(self, f, callback, pw_id, *args, **kwargs):
        self.f = f
        self.callback = callback
        self.pw_id = pw_id
        self.args = args
        self.kwargs = kwargs


class PowerGrid:
    def __init__(self, _num_threads):
        self._num_threads = _num_threads
        self._pw_objects = Queue()
        self._threads = []
        self._tasks = [Queue() for _ in range(_num_threads)]
        self._results = Queue()

    def create_pw_pool(self):
        for i in range(self._num_threads):
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
            # Store the stream in a queue, along with the ID of the stream
            self._pw_objects.put((i, pw_stream))

        for i in range(self._num_threads):
            self._threads.append(PowerThread(self._tasks, self._results, self._pw_objects))

    def threaded(self, f, daemon=False):
        def wrapped(*args, **kwargs):
            results = Queue()

            self.result_queue = results

            return self.result_queue

        return wrapped

    def reset(self):
        for i in range(self._num_threads):
            # Clean COM objects
            pw = self._pw_objects.get()[1]
            pythoncom.CoReleaseMarshalData(pw)
            pw = None
            self._pw_objects.task_done()
            # Clean task queue
            self.tasks[i].queue.clear()
            self.tasks[i].all_tasks_done.notify_all()
        self._pw_objects = None
        self._tasks = None
        # The queue should be empty at this point, this is a fallback to forcefully clear the queue
        # Doesn't kill COM objects though
        # self._pw_objects.mutex.acquire()
        # self._pw_objects.queue.clear()
        # self._pw_objects.all_tasks_done.notify_all()
        # self._pw_objects.unfinished_tasks = 0
        # self._pw_objects.mutex.release()
