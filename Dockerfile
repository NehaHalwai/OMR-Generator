# Dockerfile 
FROM python:3.11-slim

# set working dir
WORKDIR /app

# avoid writing pyc files and make logs unbuffered
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Install necessary system packages for Pillow / reportlab etc.
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    libjpeg-dev \
    zlib1g-dev \
    libfreetype6-dev \
 && rm -rf /var/lib/apt/lists/*

# copy and install python deps first for better caching
COPY requirements.txt .
RUN pip install --upgrade pip
RUN pip install --no-cache-dir -r requirements.txt

# copy app sources (images, code, assets)
COPY . .

# optional: set streamlit to be headless using CLI later; expose default port
EXPOSE 8501

# default command: run your Streamlit app (change app.py if your filename differs)
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
