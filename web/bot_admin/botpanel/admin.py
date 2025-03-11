from django.contrib import admin
from .models import BotUser, Operation, Weekend

class OperationInline(admin.TabularInline):
    model = Operation
    extra = 0
    readonly_fields = ('operation', 'created_at')
    can_delete = False

class WeekendInline(admin.TabularInline):
    model = Weekend
    extra = 0
    readonly_fields = ('date',)
    can_delete = False

@admin.register(BotUser)
class BotUserAdmin(admin.ModelAdmin):
    list_display = ('telegram_id', 'full_name', 'department', 'position', 'is_admin', 'current_shift_active')
    search_fields = ('telegram_id', 'full_name', 'department', 'position')
    list_filter = ('department', 'is_admin')
    inlines = [OperationInline, WeekendInline]


@admin.register(Operation)
class OperationAdmin(admin.ModelAdmin):
    list_display = ('user', 'operation', 'created_at')
    list_filter = ('operation', 'created_at', 'user')
    search_fields = ('user__telegram_id', 'user__full_name', 'operation')


@admin.register(Weekend)
class WeekendAdmin(admin.ModelAdmin):
    list_display = ('user', 'date')
    list_filter = ('date', 'user')
    search_fields = ('user__telegram_id', 'user__full_name')
