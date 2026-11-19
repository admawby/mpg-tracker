FROM python:3.12-slim
RUN pip install --no-cache-dir openpyxl==3.1.5
WORKDIR /app
COPY app.py /app/app.py
COPY static /app/static
ENV PYTHONUNBUFFERED=1 MPG_DB=/data/mpg.sqlite PORT=8840
USER 1000:1000
EXPOSE 8840
HEALTHCHECK --interval=60s --timeout=5s CMD python -c "import urllib.request;urllib.request.urlopen('http://127.0.0.1:8840/api/health')"
CMD ["python", "app.py"]
