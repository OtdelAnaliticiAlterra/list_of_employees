import time
from datetime import datetime
from education import education
from employees import employees
from work_exp import work_exp
from competence import competence
from photo import photos
from profile import profile
from rec import rec

while True:
    # print('Старт:', datetime.now())
    # employees()
    # education()
    # work_exp()
    # photos()
    # profile()
    # rec()
    # competence()
    #
    # print('Завершено:', datetime.now())
    # break
    if datetime.now().hour == 7:
        print('Старт:', datetime.now())
        employees()
        education()
        work_exp()
        photos()
        profile()
        rec()
        competence()
        print('Завершено:', datetime.now())
        time.sleep(60 * 60 * 12)
    else:
        time.sleep(60 * 30)