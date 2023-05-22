lint:
	pflake8 .
	isort --check .
	black --check .

format:
	isort .
	black .

type-check:
	mypy .

all:
	make format
	make lint
	make type-check
