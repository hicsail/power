import win32com.client
from win32com.client import VARIANT

import pythoncom
from queue import Queue

class PowerGrid:
    def __init__(self, _num_threads):
        self._num_threads = _num_threads
        self.results = Queue()
        self._pw_objects = Queue()

        # self._create_pw_pool(self.threads)

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

    def threaded(self, f, daemon=False):
        from threading import Thread

        def threaded_f(results, *args, **kwargs):
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

            # Call the actual threaded function
            com_result = f(*args, com_id=pw_id, auto_sim=pw, **kwargs)
            # Return function results through Queue
            results.put(com_result)

            # Revert stream back to start position
            pw_stream.Seek(0, 0)
            # Return stream back to queue
            self._pw_objects.put((pw_id, pw_stream))
            # Indicate all work on this queue object is done. Without this the queue task counter would go up every time
            # a stream is re-added to the queue
            self._pw_objects.task_done()
            # Clean up COM reference
            pw = None
            # Indicate that no more COM objects will be called in this thread
            pythoncom.CoUninitialize()

        def wrapped(*args, **kwargs):
            results = Queue()

            threads = [Thread(target=threaded_f, args=(results,)+args, kwargs=kwargs) for i in range(self._num_threads)]
            for t in threads:
                t.daemon = daemon
            [t.start() for t in threads]
            [t.join() for t in threads]
            self.result_queue = results

            return self.result_queue

        return wrapped

    def kill_com_objects(self):
        for i in range(self._num_threads):
            pw = self._pw_objects.get()[1]
            pythoncom.CoReleaseMarshalData(pw)
            pw = None
            self._pw_objects.task_done()
        self._pw_objects = None
        # The queue should be empty at this point, this is a fallback to clear the queue
        # Doesn't kill COM objects though
        # self._pw_objects.mutex.acquire()
        # self._pw_objects.queue.clear()
        # self._pw_objects.all_tasks_done.notify_all()
        # self._pw_objects.unfinished_tasks = 0
        # self._pw_objects.mutex.release()
