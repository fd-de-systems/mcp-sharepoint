FROM python:3.11-slim
WORKDIR /app

# System deps for PyMuPDF (PDF text extraction) & build tools
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential libglib2.0-0 libxrender1 libxext6 libsm6 \
 && rm -rf /var/lib/apt/lists/*

# Copy project metadata and install
COPY pyproject.toml ./
RUN pip install --no-cache-dir --upgrade pip && pip install --no-cache-dir .

# Copy source code
COPY src/ ./src/

# Container port (Azure uses this)
ENV PORT=8000

# Start the MCP server; repo uses the module entry-point
CMD ["python", "-m", "mcp_sharepoint", "--host", "0.0.0.0", "--port", "8000"]
