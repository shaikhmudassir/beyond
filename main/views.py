from django.http import HttpResponseRedirect, HttpResponse, Http404
from django.contrib.auth import authenticate
from django.shortcuts import render
from django.conf import settings
from django  import forms
from django.views import View
from beyond.settings import BASE_DIR
from .models import *
from io import BytesIO
import glob
import zipfile
from . import extract
from . import email
import os

class upload(forms.Form):
    file = forms.FileField(widget=forms.ClearableFileInput(attrs={'multiple': True}))

class Login(View):
    def get(self, request):
        if 'user_id' in request.session:
            return HttpResponseRedirect('/')
        
        return render(request, 'main/login.html', {
        })

    def post(self, request, *args, **kwargs):
        email = request.POST['email']
        password = request.POST['password']
        user = authenticate(username=email.split('@')[0], password=password)
        if user is not None:
            request.session['user_id'] = user.id
            return HttpResponseRedirect('/')
        else:
            return HttpResponseRedirect('login')

class Index(View):
    def get(self, request):
        if 'user_id' not in request.session:
            return HttpResponseRedirect('login')
        form = upload()
        return render(request, 'main/index.html', {'form':form})

class Email(View):
    def get(self, request):
        if 'user_id' not in request.session:
            return HttpResponseRedirect('login')
        #email.main()

        byte_data = BytesIO()
        zip_name = "emails.zip"
        zip_file = zipfile.ZipFile(byte_data, 'w')
        filelist = glob.glob(str(BASE_DIR) + '/main/static/main/input/*.pdf')
        for file in filelist:
            filename = os.path.basename(os.path.normpath(file))
            zip_file.write(file, filename)
        zip_file.close()

        response = HttpResponse(byte_data.getvalue(), content_type='application/zip')
        response['Content-Disposition'] = 'attachment; filename=%s' %zip_name

        zip_file.printdir()

        return response
        
        form = upload()
        return render(request, 'main/index.html', {'form':form})

class Extract(View):
    def get(self, request):
        if 'user_id' not in request.session:
            return HttpResponseRedirect('login')
        extract.extract_all_pdfs(str(BASE_DIR) + '/main/static/main/input/',str(BASE_DIR) + '/main/static/main/output/')

        file_path = str(BASE_DIR) + '/main/static/main/output/*.pdf'
        if os.path.exists(file_path):
            with open(file_path, 'rb') as fh:
                response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel")
                response['Content-Disposition'] = 'inline; filename=' + os.path.basename(file_path)
                return response
        raise Http404
        
        form = upload()
        return render(request, 'main/index.html', {'form':form})

# class Download(View):
#     def get(self, request):
#         file_path = str(BASE_DIR) + '/main/static/main/output/Report3.xlsx'
#         if os.path.exists(file_path):
#             with open(file_path, 'rb') as fh:
#                 response = HttpResponse(fh.read(), content_type="application/vnd.ms-excel")
#                 response['Content-Disposition'] = 'inline; filename=' + os.path.basename(file_path)
#                 return response
#         raise Http404

class Upload(View):
    def post(self, request):
        form = upload(request.POST,request.FILES)
        if form.is_valid():
            for f in request.FILES.getlist('file'):
                self.handle_uploaded_file(f)
            return render(request,'main/success.html',{})
        else:
            form = upload()
            return render(request,'main/index.html',{'form':form})

    def get(self, request):
        form = upload()
        return render(request,'main/index.html',{'form':form})

    def handle_uploaded_file(self,f):
        with open(str(BASE_DIR) + '/main/static/main/input/' + f.name,'wb+') as destination:
            for chunk in f.chunks():
                destination.write(chunk)

class Logout(View):
    def get(self, request):
        try:
            del request.session['user_id']
            return HttpResponseRedirect('login')
        except KeyError as e:
            return HttpResponseRedirect('/')