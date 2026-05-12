"""
Hostinger Python App entry point.
Hostinger uses Phusion Passenger which expects a WSGI application named `application`.
FastAPI is ASGI, so we wrap it via a2wsgi.
"""
from a2wsgi import ASGIMiddleware
from server import app as fastapi_app

application = ASGIMiddleware(fastapi_app)
