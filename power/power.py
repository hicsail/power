import win32com.client
from win32com.client import VARIANT

import sys
import pythoncom
import queue
from queue import Queue
from threading import Thread
from typing import Sequence, List, Callable


class PowerThread(Thread):
    """
    A thread that runs indefinitely and uses a queue to obtain more tasks.

    :type _thread_id: int
    :type _results: Queue
    :type _pw_objects: list
    :type _tasks_list: list[Queue]
    :type _tasks: Queue
    :type _pw: CDispatch
    :type _pw_stream: PyIStream
    :type _dismissed: bool
    """
    def __init__(self, i: int, task_list: Sequence[Queue], results: Queue, pw_objects: list, **kwargs):
        Thread.__init__(self, **kwargs)
        self.setDaemon(0)
        self._thread_id = i
        self._results = results
        self._pw_objects = pw_objects
        self._tasks = task_list[i]
        self._pw, self._pw_stream = None, None
        self._dismissed = False
        self.start()

    def marshal_com(self):
        """
        Marshal the PowerWorld COM object to be used in this thread.
        :return: Tuple with ID of COM object, the COM object itself and the stream referencing the COM object.
        """
        # Enable COM object access in this thread, but not others
        pythoncom.CoInitialize()
        # Get stream reference from list of com objects
        pw_stream = self._pw_objects[self._thread_id]
        # Make sure we're at the start of the stream, reset the pointer
        pw_stream.Seek(0, 0)
        # Unmarshal the stream, going back to the original interface
        pw_interface = pythoncom.CoUnmarshalInterface(pw_stream, pythoncom.IID_IDispatch)
        # And finally return the COM object that was created earlier
        pw = win32com.client.Dispatch(pw_interface)

        return pw, pw_stream

    def unmarshal_com(self):
        """Unmarshal the PowerWorld COM object and return it to the queue for later usage."""
        # Revert stream back to start position
        self._pw_stream.Seek(0, 0)
        # Clean up COM reference
        self._pw = None
        # Indicate that no more COM objects will be called in this thread
        pythoncom.CoUninitialize()

    def run(self):
        """Continuously run thread, consuming new tasks as we go."""
        # Can't do this before the thread starts running, otherwise marshalling is not successful
        self._pw, self._pw_stream = self.marshal_com()
        while True:
            if self._dismissed:
                # If there were any tasks left, get rid of them
                self._tasks.queue.clear()
                self._tasks.all_tasks_done.notify_all()
                self._tasks.unfinished_tasks = 0
                self.unmarshal_com()
                break
            try:
                # Get task with non-blocking Queue call
                task = self._tasks.get(False)
            except queue.Empty:
                continue
            else:
                try:
                    # Call task function and store results
                    result = task.f(*task.args, thread_id=self._thread_id, auto_sim=self._pw, **task.kwargs)
                    self._tasks.task_done()
                    self._results.put((task, result))
                except:
                    # Or store exception message if something went wrong
                    self._results.put((task, sys.exc_info()))

    def dismiss(self):
        """Stop executing tasks and let the thread exit."""
        self._dismissed = True


class PowerTask:
    def __init__(self, f, thread_id, *args, **kwargs):
        self.f = f
        self.thread_id = thread_id
        self.args = args
        self.kwargs = kwargs


class Power:
    """
    :type _num_threads: int
    :type _pw_objects: list
    :type _threads: list[PowerThread]
    :type _dismissed_threads: list[PowerThread]
    :type _tasks: list[Queue]
    :type _results: Queue
    """
    def __init__(self, _num_threads):
        self._num_threads = _num_threads
        self._pw_objects = []
        self._threads = []
        self._dismissed_threads = []
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
            # Store the stream in a list
            self._pw_objects.append(pw_stream)

        for i in range(self._num_threads):
            self._threads.append(PowerThread(i, self._tasks, self._results, self._pw_objects))

    def add_task(self, f: Callable, threads: str, *args, **kwargs):
        requests = set()
        for i in self._parse_thread_list(threads):
            self._tasks[i].put(PowerTask(f, i, *args, **kwargs))
            requests.add(i)

        results = []
        while True:
            if not requests:
                break
            task, result = self._results.get(True)
            results.append(result)
            requests.remove(task.thread_id)

        return results

    def dismiss_threads(self, threads: str):
        for i in self._parse_thread_list(threads):
            self._threads[i].dismiss()
            self._dismissed_threads.append(self._threads[i])
        for thread in self._dismissed_threads:
            thread.join()
        self._dismissed_threads = []

    def reset(self):
        for i in range(self._num_threads):
            # Clean COM objects
            pw = self._pw_objects[i]
            pythoncom.CoReleaseMarshalData(pw)
            pw = None
        self._pw_objects = None
        self._tasks = None
        # TODO join threads first?
        self._threads = None

        # The queue should be empty at this point, this is a fallback to forcefully clear the queue
        # Doesn't kill COM objects though
        # self._pw_objects.mutex.acquire()
        # self._pw_objects.queue.clear()
        # self._pw_objects.all_tasks_done.notify_all()
        # self._pw_objects.unfinished_tasks = 0
        # self._pw_objects.mutex.release()

    @staticmethod
    def _parse_thread_list(threads: str) -> List:
        """
        Parse string of comma and dash separated values into list that contains individual indexes.
        For example, '1,2,4-6' becomes [1,2,4,5,6]
        :param threads: String with comma and dash separated values
        :return: List with indexes of threads
        """
        # See https://stackoverflow.com/a/5705014/4573362
        ranges = (x.split("-") for x in threads.split(","))
        return [i for r in ranges for i in range(int(r[0]), int(r[-1]) + 1)]
