from typing import Any, Dict, Iterator, List
import pandas as pd
from itertools import dropwhile
from datetime import time
import re
import sys
import getopt

def time_from_str(str):
    result = re.match('(\\d+):(\\d+):(\\d+)', str)
    if result:
        hour, minute, second = tuple(map(lambda x: int(x), result.group(1, 2, 3)))
        return time(hour=hour, minute=minute, second=second)
    else:
        return time()

def seconds_of_time(time: time):
    return time.hour * 3600 + time.minute * 60 + time.second
    
def main(argv):
    try:                                                                      
        options, args = getopt.getopt(argv, "hi:m:r:d:", ["help", "input=", "members=", "ratio=", 'duration='])
    except getopt.GetoptError:
        print('使用 "-h" 查看帮助。')
        sys.exit()                                                                

    time_ratio = 0
    minimum_duration = None
    input_file = 'report.xlsx'
    member_file = 'members.txt'
    for option, value in options:
        if option in ('-h', '--help'):
            print('-i：腾讯会议导出的 xls/xlsx 文件（默认为 report.xlsx）\n-m：成员名单（默认为 members.txt）\n-r：出勤所需的参会时间占主讲人参会时间的比例（默认为 0.0 ，即有进入会议的都算出勤）\n-d：出勤所需的最少时长（可选）\n当出勤所需的最少时长被指定时，将不会按时间比例指定是否出勤。')
            return
        elif option in ('-i', '--input'):
            input_file = value
        elif option in ('-m', '--member'):
            member_file = value
        elif option in ('-r', '--ratio'):
            time_ratio = float(value)
        elif option in ('-d', '--duration'):
            minimum_duration = time_from_str(value)

    attendences: Dict[str, bool] = dict(map(lambda x: (x, False), filter(lambda x: len(x), map(lambda x: x.strip(), open(member_file, 'r').readlines()))))
    data = pd.read_excel(input_file)
    data_list = data.values.tolist()

    iter_rows: Iterator[List[Any]] = dropwhile(lambda x: not isinstance(x[0], str) or '用户名称' not in x[0], data_list)

    next(iter_rows)

    names_with_duration = map(lambda x: (x[0], x[4], x[5]), iter_rows)

    durations: List[tuple[str, int]] = []

    total_seconds = -1

    for name, duration, iden in names_with_duration:
        if isinstance(duration, str):
            duration = time_from_str(duration)

        seconds = seconds_of_time(duration)

        if '主持人' in iden:
            total_seconds = seconds
        else:
            durations.append((name, seconds))

    for name_in_meeting, seconds in durations:
        for name in attendences.keys():
            if name in name_in_meeting:
                attendences[name] = minimum_duration and seconds >= seconds_of_time(minimum_duration) or seconds >= int(time_ratio * total_seconds)

    print('缺勤人员如下：')

    for name, _ in filter(lambda x: not x[1], attendences.items()):
        print(name)

if __name__ == "__main__":
    main(sys.argv[1:])
