from django.urls import path

from .views import *

urlpatterns = [
    path('convert', ExportToXl.as_view(), name="convert_to_excel"),
]
