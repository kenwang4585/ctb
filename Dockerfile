FROM python:3

MAINTAINER Kenwang

WORKDIR /ctb

COPY Pipfile ./
COPY Pipfile.lock ./
COPY gunicorn.conf ./
COPY .env ./
RUN pip install pipenv
RUN pipenv install --system --deploy --ignore-pipfile

COPY . .

ENTRYPOINT ["./gunicorn.sh"]

#CMD ["gunicorn", "wsgi:app", "-c", "gunicorn.conf"]

#CMD [ "python", "./wsgi.py" ]