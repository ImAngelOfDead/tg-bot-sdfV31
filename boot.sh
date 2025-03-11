#!/bin/bash



echo "Запускаем сервер Django..."
python /web/bot_admin/manage.py runserver &

echo "Запускаем бота..."
python bot.py

