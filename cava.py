import time

import datetime

timenow = datetime.datetime.now()
timemidnight = datetime.datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)+datetime.timedelta(days=1)

print((timemidnight-timenow).total_seconds())