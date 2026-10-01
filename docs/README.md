# XD Allocation Hub Documentation

Welcome to the documentation suite for the **XD Allocation Hub**.

This directory contains technical documentation, architecture blueprints, change logs, and log & troubleshooting guides for the system.

## Documentation Index

- [Architecture & Data Pipeline Guide](architecture.md): System design, component breakdown, city data mappings, and processing flow.
- [UniOps & Graft Integration Guide](../uniops/README.md): Centralized MCP server, dynamic config scaling, and TrailHQ Graft context graph.
- [Changes Log & History](changes_log.md): Record of updates, features, and parameter adjustments (including the Bengaluru Staples `OU83946700` update and Vercel Web Analytics integration).
- [Logs & Troubleshooting Guide](logs_and_troubleshooting.md): Operational log management, error diagnostics, debugging steps, and Vercel deployment procedures.

## Quick Overview

The **XD Allocation Hub** is a Python/Flask web application designed for deployment on **Vercel Serverless Functions**. It processes daily e-commerce allocation indents, extra add-ons (e.g. Pluckk), and Staples uploads across 6 active cities:

1. **Chennai** (`MAA`)
2. **Trichy** (`TRZ`)
3. **Coimbatore** (`CJB`)
4. **Bengaluru** (`BLR`)
5. **Mumbai** (`BOM`)
6. **Jaipur** (`JAI`)

The application formats, coalesces, calculates allocations, assigns city/vertical contracts & supplier OU IDs, and generates multi-sheet Excel reports and zip packages for operations.
