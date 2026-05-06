FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV HOST=0.0.0.0
ENV PORT=8787

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends fonts-ipafont-gothic \
    && rm -rf /var/lib/apt/lists/*

COPY gas/python_pdf_api/requirements.txt ./requirements.txt
RUN pip install --no-cache-dir -r requirements.txt

COPY gas/python_pdf_api ./gas/python_pdf_api
COPY *.pdf ./

RUN mkdir -p gas/python_pdf_api/outputs

EXPOSE 8787

CMD ["python", "gas/python_pdf_api/server.py"]
