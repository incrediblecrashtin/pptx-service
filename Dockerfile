FROM nikolaik/python-nodejs:python3.11-nodejs20

WORKDIR /app

COPY requirements.txt .
RUN pip install -r requirements.txt

COPY package.json .
RUN npm install

COPY . .

EXPOSE 5000

CMD ["gunicorn", "app:app", "--bind", "0.0.0.0:5000"]
