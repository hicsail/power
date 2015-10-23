import win32com.client
from win32com.client import VARIANT

import sys
import pythoncom
import queue
from queue import Queue
from threading import Thread
from threading import Lock
from typing import Sequence, List, Callable


class Power:
    """
    The main class to use

    :param num_threads: Integer, amount of threads and COM objects you want to create.
    :type _num_threads: int
    :type _pw_objects: list
    :type _threads: list[_PowerThread]
    :type _dismissed_threads: list[_PowerThread]
    :type _tasks: list[Queue]
    :type _results: Queue
    :type _lock: Lock
    """
    def __init__(self, num_threads: int):
        self._num_threads = num_threads
        self._pw_objects = []
        self._threads = []
        self._dismissed_threads = []
        self._tasks = [Queue() for _ in range(num_threads)]
        self._results = Queue()
        self._lock = Lock()

    def create_pw_collection(self):
        """
        Create the collection of COM objects, equal to the thread count defined earlier when creating the Power object.
        All COM objects correspond to a specific thread and task queue.
        You should call this before add_task()
        """
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
            self._threads.append(_PowerThread(i, self._tasks, self._results, self._pw_objects, self._lock))

    def add_task(self, f: Callable, threads=None, *args, **kwargs):
        """
        Blocking call to run a method in a number of threads. The return value of each thread will be aggregated into
        one list and returned from this method.

        :param f:   The method you want to call in a thread. To use the COM object in this thread, make sure you have a
                    named parameter auto_sim, set to None. The second named parameter available is thread_id.
        :param threads: Optional string of threads to run the method in. Follows comma and dash separated notation like
                        '0-7' or '1,2,5-7'. If not provided, default to all threads.
        :param args: Any additional parameters you want to pass along
        :param kwargs: Any additional named parameters you want to pass along
        :return:    List of result tuples. The first element is the task you created with thread_id probably the most
                    useful value. The second element is the return value of your method
        """
        # If not provided, default to all threads
        if not threads:
            threads = self._all_threads()

        # Keeps track of the amount of tasks/requests currently running, so we now when all results have come in
        requests = set()
        for i in self._parse_thread_list(threads):
            self._tasks[i].put(_PowerTask(f, i, *args, **kwargs))
            requests.add(i)

        results = []
        while True:
            # Stop trying to get results if all tasks are handled
            if not requests:
                break
            # Task is returned to get to thread_id
            # Block queue while waiting for results
            task, result = self._results.get(True)
            results.append((task, result))
            # Remove thread id from requests to keep track of running tasks
            requests.remove(task.thread_id)

        return results

    def dismiss_threads(self, threads=None):
        """
        Dismiss and join a number of threads. You will most likely want to call this method with '0-7' as the
        parameter to dismiss all threads.

        :param threads: Optional string of comma and dash separated values, for example '0-7' or '1,2,5-7', defaults
                        to all threads if not provided
        """
        # If not provided, default to all threads
        if not threads:
            threads = self._all_threads()

        for i in self._parse_thread_list(threads):
            self._threads[i].dismiss()
            self._dismissed_threads.append(self._threads[i])
        # Threads are not daemons so we get a chance to release COM apartment, needs join for this reason
        for thread in self._dismissed_threads:
            thread.join()
        self._dismissed_threads = []

    def delete_pw_collection(self):
        """
        Releases the collection of COM objects and cleans up all tasks. This method should be called after
        dismiss_threads()
        """
        for i in range(self._num_threads):
            # Clean COM objects
            pw = self._pw_objects[i]
            pythoncom.CoReleaseMarshalData(pw)
            pw = None
        self._pw_objects = None
        self._tasks = None
        # TODO force join threads first?
        self._threads = None

    def _all_threads(self):
        """
        Get string notation for range of threads in the form of '0-self._num_threads'

        :return: String representing all threads that can be parsed by _parse_thread_list()
        """
        return '0' if self._num_threads is 1 else '0-' + str(self._num_threads - 1)

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


class _PowerThread(Thread):
    """
    A thread that runs indefinitely and uses a queue to obtain more tasks.

    :type _thread_id: int
    :type _results: Queue
    :type _pw_objects: list
    :type _lock: Lock
    :type _tasks_list: list[Queue]
    :type _tasks: Queue
    :type _pw: CDispatch
    :type _pw_stream: PyIStream
    :type _dismissed: bool
    """
    def __init__(self, i: int, task_list: Sequence[Queue], results: Queue, pw_objects: list, lock: Lock, **kwargs):
        Thread.__init__(self, **kwargs)
        self.setDaemon(0)
        self._thread_id = i
        self._results = results
        self._pw_objects = pw_objects
        self._lock = lock
        self._tasks = task_list[i]
        self._pw, self._pw_stream = None, None
        self._dismissed = False
        self.start()

    def marshal_com(self):
        """
        Marshal the PowerWorld COM object to be used in this thread.
        :return: Tuple with the COM object itself and the stream referencing the COM object.
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
            # Check if we're not trying to kill the thread
            if self._dismissed:
                # If not using a lock here, the main thread will throw an exception when calling reset()
                self._lock.acquire()
                try:
                    self.unmarshal_com()
                    # If there were any tasks left, get rid of them
                    self._tasks.queue.clear()
                    self._tasks.all_tasks_done.notify_all()
                    self._tasks.unfinished_tasks = 0
                finally:
                    self._lock.release()
                    break
            try:
                # Get task with blocking queue call
                # This is much cheaper than running the while loop constantly. The timeout is meant for dismissing
                # the thread, otherwise that check would never happen.
                task = self._tasks.get(True, 1)
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


class _PowerTask:
    """
    A task to be executed by a thread. We don't pass functions directly to the thread but have this intermediate
    structure to have the option of gaining more information later on. You can see which thread was responsible for
    example.

    :type f: Callable
    :type thread_id: int
    """
    def __init__(self, f: Callable, thread_id: int, *args, **kwargs):
        self.f = f
        self.thread_id = thread_id
        self.args = args
        self.kwargs = kwargs
