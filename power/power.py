import win32com.client

import sys
import traceback
import pythoncom
import queue
from queue import Queue
from threading import Thread
from threading import Lock
from typing import Sequence, List, Callable
import gevent
from power.com import PowerSocketServer


class Power:
    """
    Power provides a multithreaded PowerWorld Simulator workflow.

    Usage:

    Initiate with 4 threads
    >>> pw = Power(4)

    Create PowerWorld COM objects and threads
    >>> pw.create_pw_collection()

    Call method in all 4 threads, blocks until all results have arrived
    >>> results = pw.add_task(threaded_func)
    Call method in thread 0, 1 and 3, also blocking
    >>> results = pw.add_task(threaded_func, threads='0-1,3')

    You can pass as many arguments to the threaded function as you like
    Since the threads property is optional, you need to have None as the second parameter
    >>> results = pw.add_task(threaded_func, None, 'test', foo='bar')

    Or alternatively name your arguments for clarity, but then all should be named
    >>> results = pw.add_task(f=threaded_func, threads=None, foo='bar')

    Results is a list of tuples: first element is task, second is result.
    Task has thread_id and exception property
    >>> print(results[0][1])

    Kill all threads and COM object
    >>> pw.reset()

    Both thread_id and auto_sim have to present at all times
    >>> def threaded_func(thread_id, auto_sim):
    >>>    print('Starting thread %s' % thread_id)

    Or with arguments
    >>> results = pw.add_task(threaded_func, None, 'foo', bar='bar')
    Unnamed arguments go in front, named ones can be anywhere else
    >>> def threaded_func(foo, thread_id, auto_sim, bar):
    >>>    print('Starting thread %s' % thread_id)


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
        if num_threads < 1:
            raise ValueError('Power should be instantiated with at least 1 thread')
        self._num_threads = num_threads
        self._pw_objects = []
        self._threads = []
        self._dismissed_threads = []
        self._tasks = [Queue() for _ in range(num_threads)] # _ is a throwaway variable name that isn't used elsewhere
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

    def add_task(self, f: Callable, threads: str=None, *args, **kwargs):
        """
        Blocking call to run a method in a number of threads. The return value of each thread will be aggregated into
        one list and returned from this method.

        Keyword arguments:
        :param f: The method you want to call in a thread. To use the COM object in this thread, make sure you have a
            named parameter simAuto, set to None. The second named parameter available is thread_id.
        :param threads: Optional string of threads to run the method in. Follows comma and dash separated notation like
            '0-7' or '1,2,5-7'. If not provided, default to all threads.
        args -- Any additional parameters you want to pass along
        kwargs -- Any additional named parameters you want to pass along

        :rtype: list[(_PowerTask, T)]
        :return: List of result tuples. The first element is the task you created, containing the thread_id and
            exception flag. The second element is the return value of your method
        """
        # Yield control back to PowerSocketServer to handle any new incoming messages
        # Has to be bigger than 0 since that's not enough time for context switching
        gevent.sleep(0.1)
        # Block when application is paused
        with PowerSocketServer.sem:
            # If not provided, default to all threads
            if not threads:
                threads = self._all_threads()

            # Keeps track of the amount of tasks/requests currently running, so we know when all results have come in
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

    def reset(self):
        """
        Cleanup all data: kills threads, clears tasks and releases COM references
        """
        # If not provided, default to all threads
        for i in range(self._num_threads):
            self._threads[i].dismiss()
            self._dismissed_threads.append(self._threads[i])
        # Threads are not daemons so we get a chance to release COM apartment, needs join for this reason
        for thread in self._dismissed_threads:
            thread.join()
        self._dismissed_threads = []
        self._tasks = None
        self._threads = None

        # Delete COM object references
        for i in range(self._num_threads):
            # Clean COM objects
            pw = self._pw_objects[i]
            pythoncom.CoReleaseMarshalData(pw)
            pw = None
        self._pw_objects = None

    def _all_threads(self):
        """
        Get string notation for range of threads in the form of '0-self._num_threads'

        :rtype: str
        :return: String representing all threads that can be parsed by _parse_thread_list()
        """
        return '0' if self._num_threads is 1 else '0-' + str(self._num_threads - 1)

    @staticmethod
    def _parse_thread_list(threads: str) -> List:
        """
        Parse string of comma and dash separated values into list that contains individual indexes.
        For example, '1,2,4-6' becomes [1,2,4,5,6]

        :param threads: String with comma and dash separated values
        :rtype: list
        :return: List with indexes of threads
        """
        # See https://stackoverflow.com/a/5705014/4573362
        ranges = (x.split("-") for x in threads.split(","))
        return set([i for r in ranges for i in range(int(r[0]), int(r[-1]) + 1)])


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
        self.daemon = False
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
                    print(sys.exc_info())
                    print(traceback.print_exc())
                    task.exception = True
                    self._results.put((task, sys.exc_info()))

    def dismiss(self):
        """Stop executing tasks and let the thread exit."""
        self._dismissed = True


class _PowerTask:
    """
    A task to be executed by a thread. We don't pass functions directly to the thread but have this intermediate
    structure to have the option of gaining more information later on. You can see which thread was responsible for
    example.

    Properties:
    f -- The method to execute in a thread
    thread_id -- The ID of the thread that will execute this task
    exception -- Flag to indicate whether or not an exception happened when executing f
    args -- Any additional parameters that have been passed along
    kwargs -- Any additional named parameters that have been passed along

    :type f: Callable
    :type thread_id: int
    :type exception: bool
    """
    def __init__(self, f: Callable, thread_id: int, *args, **kwargs):
        self.f = f
        self.thread_id = thread_id
        self.exception = False
        self.args = args
        self.kwargs = kwargs
