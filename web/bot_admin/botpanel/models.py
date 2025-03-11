from django.db import models
import datetime

class BotUser(models.Model):
    full_name = models.CharField("Полное имя", max_length=255, blank=True, null=True)
    telegram_id = models.CharField("Telegram ID", max_length=255, unique=True)
    department = models.CharField("Отдел", max_length=255, blank=True, null=True)
    position = models.CharField("Должность", max_length=255, blank=True, null=True)
    is_admin = models.BooleanField("Администратор", default=False)
    reminder = models.CharField("Напоминание", max_length=255, blank=True, null=True)

    def __str__(self):
        return self.full_name or self.telegram_id

    def current_shift_active(self):
        """
        Определяет, активна ли смена у пользователя.
        Логика: если зафиксирована операция "start_shift", а следующей операции "end_shift"
        после неё нет – считаем, что смена активна.
        """
        start_op = self.operations.filter(operation="start_shift").order_by('-created_at').first()
        if not start_op:
            return False
        end_op = self.operations.filter(operation="end_shift", created_at__gt=start_op.created_at).first()
        return end_op is None

    current_shift_active.boolean = True  # для красивого отображения галочкой в админке

    class Meta:
        db_table = "users"


class Operation(models.Model):
    OPERATION_CHOICES = (
        ("start_shift", "Начало смены"),
        ("end_shift", "Завершение смены"),
        ("start_break", "Начало перерыва"),
        ("end_break", "Завершение перерыва"),
        ("photo_received", "Фото получено"),
    )
    user = models.ForeignKey(BotUser, on_delete=models.CASCADE, related_name='operations')
    operation = models.CharField("Операция", max_length=50, choices=OPERATION_CHOICES)
    created_at = models.DateTimeField("Время", auto_now_add=True)

    def __str__(self):
        return f"{self.user} — {self.get_operation_display()} ({self.created_at.strftime('%d.%m.%Y %H:%M:%S')})"

    class Meta:
        db_table = "operations"
        ordering = ['-created_at']


class Weekend(models.Model):
    user = models.ForeignKey(BotUser, on_delete=models.CASCADE, related_name='weekends')
    date = models.DateField("Дата выходного")

    def __str__(self):
        return f"{self.user} — {self.date.strftime('%d.%m.%Y')}"

    class Meta:
        db_table = "weekends"
        ordering = ['date']
