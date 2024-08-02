# Use the official Python image from the Docker Hub
FROM python:3.11-slim

# Set the working directory in the container
WORKDIR /app

# Copy the requirements file into the container
COPY requirements.txt requirements.txt

# Install dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Copy the rest of the application code into the container
COPY . .

ENV FLASK_APP=main.py

# Expose the port Flask will run on
EXPOSE 5000

# Define the command to run the application
CMD ["python3", "main.py", "--host=0.0.0.0"]
