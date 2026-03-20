import sys
import os

# Add the src directory to sys.path so tests can import graph_auth and graph_client
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "src"))
