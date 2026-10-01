# Logs & Troubleshooting Guide

## Logging Architecture

The **XD Allocation Hub** produces server-side logs during Flask execution, database metadata queries, file processing, and report building.

- **Vercel Serverless Logs**: Operational logs (stdout/stderr) are streamed directly to the Vercel Runtime Logs dashboard when deployed on Vercel.
- **Local Logs**: When running locally (`python xd_processor.py`), logs output directly to the standard console stdout.

## Common Operational Issues & Diagnostics

### 1. Missing Supplier ID / OUID on Staples Items
- **Symptom**: Items uploaded via Staples file display default OU IDs (`OU83946715` / `OU77187305`) instead of `OU83946700`.
- **Cause**: The city selected was not `Bengaluru`, or the file was uploaded into the "Extra Add-ons" field instead of the "Staples File" field.
- **Solution**: Ensure the file is uploaded using the dedicated "Staples File" input and the "Bengaluru" card is selected.

### 2. Vercel Web Analytics Script 404 in Local Development
- **Symptom**: Console browser warning `GET http://localhost:5000/_vercel/insights/script.js 404 (Not Found)`.
- **Cause**: Vercel Web Analytics script (`/_vercel/insights/script.js`) is served automatically by Vercel edge runtime in deployed production environments.
- **Solution**: This warning is expected when testing locally and does not affect application functionality. It activates automatically when deployed to Vercel.

### 3. Missing Date Error (400 Bad Request)
- **Symptom**: API returns `{"error": "Missing 'date' field. Send { \"date\": \"YYYY-MM-DD\" }"}`.
- **Solution**: Select a date in the date picker UI or pass `date=YYYY-MM-DD` as a form field / URL parameter.

### 4. MySQL Metadata Fetch Failure
- **Symptom**: Log warning `[DB Error] Unable to connect to metadata database`.
- **Cause**: Invalid database credentials in `db_credentials.json` or network restriction.
- **Solution**: The system gracefully falls back to default vertical ("Staples") and default tag ("chiller") when database access is unavailable.

## Vercel Deployment & Verification Steps

1. **Local Syntax Verification**:
   ```bash
   python -m py_compile xd_processor.py
   ```
2. **Deploying to Vercel**:
   ```bash
   vercel --prod
   ```
3. **Checking Runtime Analytics**:
   Navigate to [Vercel Analytics Dashboard](https://vercel.com/analytics) to monitor page views, web vitals, and request metrics.
