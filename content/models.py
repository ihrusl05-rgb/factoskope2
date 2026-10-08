from django.db import models


class ZodiacSign(models.TextChoices):
    ARIES = 'aries', 'Овен'
    TAURUS = 'taurus', 'Телец'
    GEMINI = 'gemini', 'Близнецы'
    CANCER = 'cancer', 'Рак'
    LEO = 'leo', 'Лев'
    VIRGO = 'virgo', 'Дева'
    LIBRA = 'libra', 'Весы'
    SCORPIO = 'scorpio', 'Скорпион'
    SAGITTARIUS = 'sagittarius', 'Стрелец'
    CAPRICORN = 'capricorn', 'Козерог'
    AQUARIUS = 'aquarius', 'Водолей'
    PISCES = 'pisces', 'Рыбы'


class Fact(models.Model):
    text = models.TextField(blank = False, verbose_name = "Текст")
    category = models.CharField(max_length=64, blank = True, verbose_name = "Категория")
    ordering = models.PositiveIntegerField (default = 0, verbose_name = "Порядок")
    is_active = models.BooleanField (default = True, verbose_name = "Активно")
    
    class Meta:
        verbose_name = 'Факт'
        verbose_name_plural = 'Факты'
        ordering = ['ordering', 'id']

    def __str__(self):
        return self.text[:30] + ('...' if len(self.text) > 30 else '') 



class Horoscope(models.Model):
    sign = models.CharField(max_length=20, choices=ZodiacSign.choices, verbose_name="Знак зодиака", db_index=True)
    date = models.DateField(verbose_name="Дата", db_index=True)
    text = models.TextField(verbose_name = "Текст")
    is_active = models.BooleanField (default = True, verbose_name = "Активно")
    is_draft = models.BooleanField (default = False, verbose_name = "Черновик")
    
    class Meta:
        verbose_name = 'Гороскоп'
        verbose_name_plural = 'Гороскопы'
        ordering = ['-date', 'sign']
        constraints = [models.UniqueConstraint(fields = ['date', 'sign'], name= "unique_sign_date")]

    def __str__(self):
        return f'{self.get_sign_display()} — {self.date.strftime("%d.%m.%Y")}'       