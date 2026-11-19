run:
	python -m flask --app app run

init-admin:
	python -m flask --app app init-admin

install:
	python -m pip install -r requirements.txt

venv:
	python -m venv .venv

mysql-up:
	docker compose up -d db

app-up:
	docker compose up -d app

all-up:
	docker compose up -d

down:
	docker compose down

clean:
	docker compose down -v
