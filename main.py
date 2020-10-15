import argparse
import psutil
import datetime
from enum import Enum
import pandas as pd
import time
import sys
import signal
import os


class Process():
    class State(Enum):
        On = 1
        Off = 2

    def __init__(self, name:str):
        self.name = name
        self.start_time = None
        self.end_time = None
        self.duration = None

    def get_pid(self) -> int:
        for proc in psutil.process_iter():
            if proc.name() == args.process:
                return proc.pid
        return -1

    def get_state(self) -> State:
        for proc in psutil.process_iter():
            if proc.name() == args.process:
                return Process.State.On
        return Process.State.Off

    def wait_until_started(self):
        state = self.get_state()
        while state == Process.State.Off:
            time.sleep(0.1)
            state = self.get_state()

    def prepare_to_monitor(self, duration:pd.Timedelta):
        self.duration = duration
        self.start_time = datetime.datetime.now()
        self.end_time = self.start_time + self.duration

    def monitor(self):
        interval = self.duration.total_seconds() / 100
        state = self.get_state()
        current_time = datetime.datetime.now()
        while state == Process.State.On and current_time < self.end_time:
            percentage = get_duration_percentage(self.start_time, self.end_time, current_time)
            print("Monitoring, ", percentage, "% of total duration passed")
            time.sleep(interval)
            current_time = datetime.datetime.now()
            state = self.get_state()

    def terminate(self):
        if sys.platform == 'win32':
            import ctypes
            PROCESS_TERMINATE = 1
            handle = ctypes.windll.kernel32.OpenProcess(PROCESS_TERMINATE, False, self.get_pid())
            ctypes.windll.kernel32.TerminateProcess(handle, -1)
            ctypes.windll.kernel32.CloseHandle(handle)
        else:
            os.kill(self.get_pid(), signal.SIGKILL)


class Report():
    def __init__(self, process:Process, report_dir: str):
        self.date = datetime.datetime.now().date()
        self.report_dir = report_dir
        self.process = process
        self.report_file_name = self._get_report_file_name()
        self.report_df = self._get_report_df()
        self.times_opened = self._get_times_opened_today()

    def _get_last_index(self, df:pd.DataFrame):
        return len(df.index) - 1

    def _get_report_df(self):
        if not os.path.exists(self.report_file_name):
            self._create_new_report_csv()
        return pd.read_csv(self.report_file_name)

    def _create_new_report_csv(self):
        columns = ['date', 'start_time', 'expected_end_time', 'actual_end_time', 'duration', 'times_opened_today']
        new_df = pd.DataFrame(columns=columns)
        new_df.to_csv(self.report_file_name, index=False, columns=columns)

    def _get_report_file_name(self) -> str:
        return self.report_dir + os.sep + self.process.name + "_report.csv"

    def _get_times_opened_today(self):
        today_df = self.report_df.loc[self.report_df["date"] == self.date.isoformat()]
        if today_df.empty:
            return 1
        else:
            last = self._get_last_index(today_df)
            times_opened = today_df.iloc[last]["times_opened_today"] + 1
            return times_opened

    def report_start(self):
        times_opened = self._get_times_opened_today()
        new_row_index = self._get_last_index(self.report_df) + 1
        new_row = {"date": self.date.isoformat(),
                   "start_time":self.process.start_time.time(),
                   "expected_end_time": self.process.end_time.time(),
                   "actual_end_time": "unknown",
                   "duration": self.process.duration,
                   "times_opened_today": times_opened}
        self.report_df.loc[new_row_index] = new_row
        self.report_df.to_csv(self.report_file_name, index=False)


    def report_termination(self):
        row_index = self._get_last_index(self.report_df)
        new_column = pd.Series([datetime.datetime.now().time().isoformat()], name="actual_end_time", index=[row_index])
        self.report_df.update(new_column)
        self.report_df.to_csv(self.report_file_name, index=False)


def get_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(formatter_class=argparse.RawTextHelpFormatter)
    parser.add_argument('--process', help='The process to monitor', required=True)
    parser.add_argument('--report-dir', help='The dir where to store reports', required=True)
    parser.add_argument('--duration', help='How long should we allow this process to be on before terminating', required=True)
    parser.add_argument('--max-times-opened', help='How many times are we allowed to open the app?', required=True)

    args = parser.parse_args()
    duration = parse_duration(args.duration)
    args.duration = duration
    return args


def get_duration_percentage(start_time: datetime.datetime,
                            end_time: datetime.datetime,
                            current_time: datetime.datetime):
    total_duration = (end_time - start_time).total_seconds()
    current_duration = (current_time - start_time).total_seconds()
    return round(100 * (current_duration / total_duration), 2)


def parse_duration(duration_string: str) -> pd.Timedelta:
    return pd.Timedelta(duration_string)


def main(args: argparse.Namespace):
    process = Process(args.process)
    report = Report(process, args.report_dir)

    print("Starting to monitor process:'", process,"'")
    state = process.get_state()

    print("Current state:", state.name)
    if state == Process.State.On:
        print("Process already on, starting duration count from now anyway...")
    else:
        print("Waiting for process to start...")
        process.wait_until_started()

    process.prepare_to_monitor(args.duration)

    report.report_start()
    if report.times_opened > int(args.max_times_opened):
        print("Application was opened too many times, (", report.times_opened, "/", args.max_times_opened, ")terminating it now")
        report.report_termination()
        process.terminate()
        return

    print("Starting monitoring process at time:", process.start_time)
    print("Process will be terminated after", process.duration.total_seconds() / 60, "minutes, at:", process.end_time)
    process.monitor()

    print("Duration passed! Terminating the process")
    report.report_termination()
    process.terminate()


if __name__ == '__main__':
    args = get_args()
    main(args)

