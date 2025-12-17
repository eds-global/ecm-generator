FROM python:3.9-slim

WORKDIR /app

# We removed software-properties-common to fix the build error
RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip3 install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8501

ENTRYPOINT ["streamlit", "run", "app1.py", "--server.port=8501", "--server.address=0.0.0.0"]