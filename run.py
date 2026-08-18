"""Entry point: run the Employment Center check-in app locally.

    python run.py            # http://127.0.0.1:5000
"""

from app.main import create_app

app = create_app()

if __name__ == "__main__":
    import os

    port = int(os.environ.get("PORT", "5000"))
    app.run(host="127.0.0.1", port=port, debug=True)
