import os
from datetime import datetime, timedelta

file_dir = "C:/Mantra/reports/nonae_copy" #location
giorni = 1 #n max of days

giorni_pass = datetime.now() - timedelta(giorni)
print(giorni_pass)
for root, dirs, files in os.walk(file_dir):

    for file in files:
        path = os.path.join(file_dir, file)
        print(path)
        filetime = datetime.fromtimestamp(os.path.getctime(path))
        print(filetime)
        if filetime > giorni_pass:
            os.remove(path)
