""" Script to test running powerfactory in multi-processing mode to allow multiple instances to be run so that """
import sys
import multiprocessing
import time
import traceback

dir_pf = r'C:\Program Files\DIgSILENT\PowerFactory 2016 SP5\Python\3.5'
sys.path.append(dir_pf)
import powerfactory


class Process(multiprocessing.Process):
    """Added to capture Exceptions that take place during multiprocessing runs"""
    def __init__(self, *args, **kwargs):
        multiprocessing.Process.__init__(self, *args, **kwargs)
        self._pconn, self._cconn = multiprocessing.Pipe()
        self._exception = None

    def run(self):
        try:
            multiprocessing.Process.run(self)
            self._cconn.send(None)
        except Exception as e:
            tb = traceback.format_exc()
            self._cconn.send((e, tb))
            # raise e  # You can still rise this exception if you need to

    @property
    def exception(self):
        if self._pconn.poll():
            self._exception = self._pconn.recv()
        return self._exception


def worker():
    app = powerfactory.GetApplication()
    time.sleep(0.5)
    app = None


if __name__=='__main__':
    jobs = []
    for i in range(3):
        job = Process(target=worker)
        job.start()
        print('Job {} started'.format(i))
        jobs.append(job)

    for job in jobs:
        job.join()
        if job.exception:
            error, traceback = job.exception
            print(traceback)
