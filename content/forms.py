from django import forms


class ExcelImportForm(forms.Form):
    """Проверяет расширение и размер Excel-файла перед чтением его строк.
    Проверка заголовков и содержимого выполняется функциями в services.py."""

    file = forms.FileField(label="Файл Excel", required=True)

    def clean_file(self):
        """Проверяет загруженный файл при вызове ``form.is_valid()``.
        Returns:
            Загруженный файл с расширением ``.xlsx`` и размером до 5 МБ.
        Raises:
            forms.ValidationError: Если расширение или размер не подходят."""
        file = self.cleaned_data.get("file")
        if file:
            if not file.name.lower().endswith(".xlsx"):
                raise forms.ValidationError("Файл должен быть в формате Excel (.xlsx).")
            if file.size > 5 * 1024 * 1024:  # 5 MB limit
                raise forms.ValidationError("Размер файла не должен превышать 5 МБ.")
        return file


class ContentImportForm(ExcelImportForm):
    """Принимает тип контента и отдельный Excel-файл для импорта."""

    content_type = forms.ChoiceField(
        label="Тип данных",
        choices=[("facts", "Факты"), ("horoscopes", "Гороскопы")],
    )
    field_order = ["content_type", "file"]
