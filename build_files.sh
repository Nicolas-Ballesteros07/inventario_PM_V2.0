#!/bin/bash
set -e
echo "Installing dependencies..."
python3 -m pip install -r requirements.txt

echo "Collecting static files..."
cd app
python3 manage.py collectstatic --noinput
python3 manage.py makemigrations
python3 manage.py migrate
cd ..