from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('upload', views.upload, name='upload'),
    path('generator', views.generator, name='generator'),
    path('download', views.download, name='download'),
    path('logout', views.logout, name='logout')
]