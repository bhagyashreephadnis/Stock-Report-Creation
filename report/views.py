from .tasks import logic
from django.shortcuts import render, redirect
from django.http import HttpResponse
from django.contrib.auth.models import User, auth
from django.contrib import messages
from datetime import date

# Create your views here.
def index(request):
    if request.method == 'POST':
        username = request.POST['username']
        password = request.POST['password']
        user = auth.authenticate(username = username, password = password)
        if user is not None:
            auth.login(request, user)
            return redirect(upload)
        else:
            messages.info(request, "Invalid Credentials")
            return redirect("/")
    else:
        return render(request, "index.html")

def upload(request):
    if request.method=='POST':
        if len(request.FILES) == 0:
            messages.info(request, "jhbj")
            # messages.add_message(request, messages.INFO, 'Upload RRP1 file')
            return redirect(upload)
        # file1 = request.FILES['pwsfile']
        file2 = request.FILES['rrp1file']
        # with open('report/input/pws.xlsx', 'wb+') as destination:
        #     for chunk in file1.chunks():
        #         destination.write(chunk)
        with open('report/input/rrp1.xlsx', 'wb+') as destination:
            for chunk in file2.chunks():
                destination.write(chunk)
        return redirect(generator)
    if request.user.is_authenticated:
        return render(request, "upload.html")
    else:
        return redirect("/")

def generator(request):
    if request.method=='POST':
        variant = request.POST.get("variant", "")
        if len(request.FILES) != 0:
            file1 = request.FILES['basefile']
            if variant=="paste":
                with open('report/input/Base File (Paste).xlsx', 'wb+') as destination:
                    for chunk in file1.chunks():
                        destination.write(chunk)
                with open('report/static/Base File (Paste).xlsx', 'wb+') as destination:
                    for chunk in file1.chunks():
                        destination.write(chunk)
                return redirect(download_paste)
            elif variant=="brush":
                with open('report/input/Base File (Brush).xlsx', 'wb+') as destination:
                    for chunk in file1.chunks():
                        destination.write(chunk)
                with open('report/static/Base File (Brush).xlsx', 'wb+') as destination:
                    for chunk in file1.chunks():
                        destination.write(chunk)
                return redirect(download_brush)
            else:
                with open('report/input/Base File (PCP).xlsx', 'wb+') as destination:
                    for chunk in file1.chunks():
                        destination.write(chunk)
                with open('report/static/Base File (PCP).xlsx', 'wb+') as destination:
                    for chunk in file1.chunks():
                        destination.write(chunk)
                return redirect(download_other)
        else:
            if variant=="paste":
                return redirect(download_paste)
            elif variant=="brush":
                return redirect(download_brush)
            else:
                return redirect(download_other)
    elif request.user.is_authenticated:
        return render(request, "generator.html")
    else:
        return redirect("/")

def download_paste(request):
    if request.user.is_authenticated:
        logic_task = logic.delay("paste")
        task_id = logic_task.task_id
        filename = date.today().strftime("%d%b%Y")+" (Paste).xlsx"
        return render(request, "download.html", {'filename':filename, 'task_id':task_id})
    else:
        return redirect("/")

def download_brush(request):
    if request.user.is_authenticated:
        logic_task = logic.delay("brush")
        task_id = logic_task.task_id
        filename = date.today().strftime("%d%b%Y")+" (Brush).xlsx"
        return render(request, "download.html", {'filename':filename, 'task_id':task_id})
    else:
        return redirect("/")

def download_other(request):
    if request.user.is_authenticated:
        logic_task = logic.delay("other")
        task_id = logic_task.task_id
        filename = date.today().strftime("%d%b%Y")+" (PCP).xlsx"
        return render(request, "download.html", {'filename':filename, 'task_id':task_id})
    else:
        return redirect("/")

def logout(request):
    auth.logout(request)
    return redirect("/")