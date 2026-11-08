# Publipostage BA22 — version web
FROM python:3.11-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    LANG=C.UTF-8

# LibreOffice (conversion DOCX -> PDF, sans interface graphique)
# + polices métriquement compatibles avec celles de Word :
#   Carlito = Calibri, Caladea = Cambria, Liberation = Arial / Times / Courier
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer-nogui \
        fonts-crosextra-carlito \
        fonts-crosextra-caladea \
        fonts-liberation2 \
        fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

# Utilisateur non root (LibreOffice a besoin d'un HOME inscriptible)
RUN useradd --create-home --uid 1000 app
WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt gunicorn

COPY --chown=app:app . .
RUN mkdir -p /app/data /home/app/.local/share/Publipostage \
    && chown -R app:app /app/data /home/app/.local

USER app
EXPOSE 8000

# Un seul worker : l'état des envois (logs en direct) reste en mémoire d'un même
# processus, et LibreOffice ne convertit qu'un document à la fois.
# Plusieurs threads pour servir les pages pendant un envoi.
# --timeout 0 : pas de coupure pendant un long envoi.
CMD ["gunicorn", "--workers", "1", "--threads", "8", "--timeout", "0", \
     "--bind", "0.0.0.0:8000", "app:app"]
