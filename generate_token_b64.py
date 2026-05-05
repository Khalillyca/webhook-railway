"""
Run this ONCE on your local machine to generate the TOKEN_CACHE_B64 value.
Then paste that value into Railway environment variables.

Usage:
    python generate_token_b64.py
"""
import os
import base64

TOKEN_CACHE_FILE = ".token_cache.bin"

if not os.path.exists(TOKEN_CACHE_FILE):
    print(f"ERROR: {TOKEN_CACHE_FILE} not found.")
    print("Run webhook_server.py locally first to generate the token cache.")
else:
    with open(TOKEN_CACHE_FILE, "r") as f:
        content = f.read()
    b64 = base64.b64encode(content.encode("utf-8")).decode("utf-8")
    print("\n" + "=" * 60)
    print("Copy this value and set it as TOKEN_CACHE_B64 in Railway:")
    print("=" * 60)
    print(b64)
    print("=" * 60 + "\n")
