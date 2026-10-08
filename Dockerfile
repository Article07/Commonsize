FROM python:3.12-slim

# Tesseract is the OCR engine used for scanned PDFs
RUN apt-get update \
    && apt-get install -y --no-install-recommends tesseract-ocr \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# One OpenMP thread per Tesseract process: with a fractional CPU quota, Tesseract's default threading makes
# OCR many times slower (threads spin and get throttled). Pages are OCR'd in parallel by the app instead,
# one process per usable CPU.
ENV OMP_THREAD_LIMIT=1
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app.py ui.py ./
COPY config ./config
COPY core ./core
COPY template ./template
COPY .streamlit ./.streamlit

# Render supplies $PORT; APP_PASSWORD is set in the Render dashboard
CMD streamlit run app.py --server.port ${PORT:-8501} --server.address 0.0.0.0
