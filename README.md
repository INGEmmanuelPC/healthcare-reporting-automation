# healthcare-reporting-automation

Automated regulatory compliance pipeline that replaced a 70-minute manual reporting process at a Colombian university hospital with a 6-minute, zero-transcription workflow. Extracts appointment wait-time data from an undocumented clinical SQL Server database, transforms it into the format required by Colombian Resolución 1552 de 2013, writes it into the official Excel template via Microsoft Graph API, and distributes the completed report by email with structured failure alerting.

This repository contains no hospital code, no patient data, and no credentials; because I sanitized my experience.

The engineering narrative document ([ENGINEERING_DECISIONS.md](./ENGINEERING_DECISIONS.md)) is the content: it walks through six failure modes I encountered while building the pipeline, each documented as symptom, diagnosis, alternatives, solution, and lesson learned.

**Stack:** n8n (self-hosted, Docker Compose/WSL2), SQL Server, Microsoft Graph API, SharePoint/OneDrive for Business, Outlook, PostgreSQL.

## What's here

| File | Description |
|---|---|
| `ENGINEERING_DECISIONS.md` | Full engineering narrative: architecture, six problems in depth, infrastructure lessons, impact metrics |

## Context

Built during a university-industry immersion (Universidad Autónoma de Manizales → hospital IT department), working alongside a cybersecurity specialist and a database engineer, advised by a PhD researcher in AI. The pipeline handles patient-identifiable data under Colombian health data regulations (Ley 1581 de 2012).
