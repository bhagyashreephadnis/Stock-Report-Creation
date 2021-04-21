from django.shortcuts import render, redirect
from django.http import HttpResponse
from django.contrib.auth.models import User, auth
from django.contrib import messages
from report.scripts.stockReport import startpoint

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
        # if len(request.FILES) == 0:
            # messages.info(request, "jhbj")
            # messages.add_message(request, messages.INFO, 'Upload both files!!!!!')
            # return render(request, "upload.html")
        file1 = request.FILES['pwsfile']
        file2 = request.FILES['rrp1file']
        with open('report/input/pws.xlsx', 'wb+') as destination:
            for chunk in file1.chunks():
                destination.write(chunk)
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
        if len(request.FILES) != 0:
            file1 = request.FILES['basefile']
            with open('report/input/Base File.xlsx', 'wb+') as destination:
                for chunk in file1.chunks():
                    destination.write(chunk)
            with open('report/static/Base File.xlsx', 'wb+') as destination:
                for chunk in file1.chunks():
                    destination.write(chunk)
        # filename = startpoint()
        return redirect(download)
    if request.user.is_authenticated:
        return render(request, "generator.html")
    else:
        return redirect("/")

def download(request):
    if request.user.is_authenticated:
        # filename = startpoint()
        filename = "14Apr2021.xlsx"
        return render(request, "download.html", {'filename':filename})
    else:
        return redirect("/")

def logout(request):
    auth.logout(request)
    return redirect("/")