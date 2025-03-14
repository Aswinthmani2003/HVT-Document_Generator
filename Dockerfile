# Use a lightweight Python image
FROM python:3.9-slim

# Set the working directory
WORKDIR /app

# Copy application files
COPY . /app

# Install system dependencies required for LibreOffice and document processing
RUN apt-get update && \
    apt-get install -y libreoffice libgl1-mesa-glx libglib2.0-0 && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

# Install Python dependencies
RUN pip install --no-cache-dir -r requirements.txt

# Expose the Streamlit app port
EXPOSE 8501

# Set the Streamlit command to run the app
CMD streamlit run app.py --server.port=$PORT --server.address=0.0.0.0