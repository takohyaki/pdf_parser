from django.urls import path
from . import views
from .views import index

urlpatterns = [
    path('upload/', views.file_upload, name='file_upload'),
    path('', index, name='index'),
]
