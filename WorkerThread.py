from PyQt5.QtCore import QThread, pyqtSignal

class WorkerThread(QThread):
    finished = pyqtSignal()
    progress = pyqtSignal(int)  # Signal to emit progress updates

    def __init__(self, task_function, *args, **kwargs):
        super().__init__()
        self.task_function = task_function
        self.args = args
        self.kwargs = kwargs

    def run(self):
        # Pass the progress callback only if the task function accepts it
        if 'progress_callback' in self.kwargs:
            self.task_function(*self.args, **self.kwargs)
        else:
            self.task_function(progress_callback=self.progress.emit, *self.args, **self.kwargs)
        self.finished.emit()