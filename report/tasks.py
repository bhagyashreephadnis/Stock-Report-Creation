from report.scripts.stockReport import startpoint
from celery.task.schedules import crontab
from celery.decorators import periodic_task, task
from celery import shared_task
from report.scripts.sapPws import pwsStart
from time import sleep

@shared_task
def run_pws():
    # print('nice')
    result = pwsStart()

@periodic_task(run_every=(crontab(hour=8, minute=30)), ignore_result=True)
def pws_task():
    run_pws.delay()

@task(name="logic", ignore_result=False)
def logic():
    # filename = "x"
    # sleep(10)
    filename = startpoint()
    return filename