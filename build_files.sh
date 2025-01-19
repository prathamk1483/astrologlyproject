echo "Deployment started"

python3.12 -m venv env

source env/bin/activate

pip install -r requirements.txt

python manage.py collectstatic --noinput

echo "Deployment completed"