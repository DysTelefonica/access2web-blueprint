"""Composition root for the lanzadera module (D73, D91, DA-1).

This package wires driven ports to adapters and exposes the
application-layer use cases as bound methods on a single
``LanzaderaContainer`` object. The HTTP delivery layer (PR 5, FastAPI
router) and the CLI driver import the container and resolve the use
cases they need; the ``bootstrap`` module runs the idempotent
admin-bootstrap step at process startup (D91 + CA-F4).
"""
