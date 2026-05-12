"""
Hostinger Python App entry point.

Hostinger's Python App uses Phusion Passenger which expects a WSGI application
named `application`. FastAPI is an ASGI app, so we wrap it via a2wsgi.

Make sure `a2wsgi` is in requirements.txt.
"""
from a2wsgi import ASGIMiddleware
from server import app as fastapi_app

application = ASGIMiddleware(fastapi_app)
