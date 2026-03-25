import os

os.system("pip install openpyxl")

import uvicorn

port = int(os.environ.get("PORT", 8000))

if __name__ == "__main__":
    uvicorn.run("app:app", host="0.0.0.0", port=port)
