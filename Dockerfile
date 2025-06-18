FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Create necessary directories
RUN mkdir -p uploads static

# Expose the port the app runs on
EXPOSE $PORT

# Command to run the application
CMD uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}
