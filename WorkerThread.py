from PyQt5.QtCore import QThread, pyqtSignal
import traceback


class WorkerThread(QThread):
    """Run a blocking task in a QThread and forward progress/result/error signals.

    Usage:
      worker = WorkerThread(task_function, *args, **kwargs)
      worker.progress.connect(lambda v: ...)
      worker.result.connect(lambda res: ...)
      worker.error.connect(lambda msg: ...)
      worker.finished.connect(lambda: ...)
      worker.start()
    """

    finished = pyqtSignal()
    progress = pyqtSignal(object)  # progress payload (int or dict)
    result = pyqtSignal(object)
    error = pyqtSignal(str)

    def __init__(self, task_function, *args, **kwargs):
        super().__init__()
        self.task_function = task_function
        self.args = args
        # copy kwargs so we can inject safely
        self.kwargs = dict(kwargs)

    def run(self):
        try:
            # Ensure the task receives a progress_callback callable
            if 'progress_callback' not in self.kwargs:
                # provide a callback that emits the `progress` signal
                self.kwargs['progress_callback'] = self.progress.emit

            # Call the task; capture return value if any
            result = self.task_function(*self.args, **self.kwargs)

            # Emit result if not None (caller can still ignore)
            try:
                self.result.emit(result)
            except TypeError:
                # signal emission may fail for unserializable objects; ignore
                pass

        except Exception as exc:
            tb = traceback.format_exc()
            # Send error details to the UI thread for logging/display
            try:
                self.error.emit(f"{exc!r}\n{tb}")
            except Exception:
                # best-effort: if emitting fails, there's nothing else we can do here
                pass
        finally:
            # Always emit finished so callers can clean up UI state
            try:
                self.finished.emit()
            except Exception:
                pass