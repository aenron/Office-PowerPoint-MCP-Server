FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    PIP_NO_CACHE_DIR=1 \
    DOWNLOAD_URL=http://localhost:8000

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        fonts-noto-cjk \
        libxml2 \
        libxslt1.1 \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt pyproject.toml README.md LICENSE ./
RUN pip install --upgrade pip \
    && pip install -r requirements.txt

COPY ppt_mcp_server.py ./
COPY tools ./tools
COPY utils ./utils
COPY templates ./templates
COPY public ./public

RUN mkdir -p /app/ppt /app/output /app/public/downloads /app/scratch

EXPOSE 8000

ENTRYPOINT ["python", "ppt_mcp_server.py"]
CMD ["--transport", "sse", "--host", "0.0.0.0", "--port", "8000"]
