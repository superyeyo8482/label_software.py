import os
import sys

# Flat, single-app layout: modules are siblings one directory up from
# tests/, not an installed package. Add that directory to sys.path so
# `import printing`, `import licensing`, etc. resolve during test collection.
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
