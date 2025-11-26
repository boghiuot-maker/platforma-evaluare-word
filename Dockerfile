FROM python:3.11-slim
WORKDIR /app
COPY . /app
RUN pip install --upgrade pip
RUN pip install -r app/requirements.txt
ENV FLASK_APP=app/server.py
EXPOSE 5000
CMD ["python", "app/server.py"]
