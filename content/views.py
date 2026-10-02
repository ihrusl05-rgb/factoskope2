from django.http import JsonResponse
from django.views import View
from .models import Horoscope, ZodiacSign, Fact
from datetime import date, timezone


class FactsAPIView(View):
    def get(self, request):
        try:
            limit = int(request.GET.get("limit", 50))
        except (TypeError, ValueError):
            limit = 50

        limit = max(0, min(limit, 100))

        facts = Fact.objects.filter(is_active=True).order_by("ordering","id",)[:limit]

        data = [
            {
                "text": fact.text,
                "category": fact.category,
            }
            for fact in facts
        ]

        return JsonResponse({"facts": data})


class HoroscopeAPIView(View):
    def get(self, request):
        today = request.GET.get("date")

        if today:
            try:
                requested_date = date.fromisoformat(today)
            except ValueError:
                return JsonResponse(
                    {"error": "invalid date, use YYYY-MM-DD"},
                    status=400,
                )
        else:
            requested_date = timezone.localdate()

        horoscopes = Horoscope.objects.filter(is_active=True, is_draft=False, date=today)

        data = {
            sign.label: horoscopes.filter(sign=sign.value).first().text if horoscopes.filter(sign=sign.value).exists() else None
            for sign in ZodiacSign
        }

        return JsonResponse({"horoscopes": data})    