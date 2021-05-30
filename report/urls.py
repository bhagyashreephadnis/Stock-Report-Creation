from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='index'),
    path('upload', views.upload, name='upload'),
    path('generator', views.generator, name='generator'),
    path('download_paste', views.download_paste, name='download_paste'),
    path('download_brush', views.download_brush, name='download_brush'),
    path('download_pcp', views.download_other, name='download_other'),
    path('logout', views.logout, name='logout')
]