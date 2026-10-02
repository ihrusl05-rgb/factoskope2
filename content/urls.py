from django.urls import path
from content import views

urlpatterns = [
    path("facts/", views.FactsAPIView.as_view(), name="facts"),
    path("horoscope/", views.HoroscopeAPIView.as_view(), name="horoscope")
]