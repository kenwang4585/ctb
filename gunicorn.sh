#!/bin/sh
gunicorn -w 2 -b 0.0.0.0:6000 wsgi:app