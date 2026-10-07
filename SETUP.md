# Setup Guide

## Install dependencies

```
pip install -r requirements.txt
playwright install chromium
```

## Configure .env

Copy `.env.example` to `.env` and fill in:

```
PROXY_POOL=http://user:pass@host:port,...
GOOGLE_API_KEY=...   (optional)
GOOGLE_CX=...         (optional)
```

## Run modes

### Standard (proxies only)

```
email-harvester --niche "plumbers" --location "Austin, TX" --max 200
```

### Browser mode (VPS recommended)

```
email-harvester --niche "plumbers" --location "Austin, TX" --browser --max 200
```

### Full power (browser + API fallback)

```
email-harvester --niche "plumbers" --location "Austin, TX" --browser --api-fallback --max 200
```

### Enrich existing sheet

```
email-harvester --enrich my_leads.xlsx --niche "plumbers" --location "Austin, TX"
```
