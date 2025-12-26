# 📦 Automated Inventory Analysis and Reporting Platform

## 🔍 Project Overview

This project is a production-ready data analytics web application developed using Python and Streamlit to analyze, visualize, and report warehouse inventory data.
It was designed to replace complex Excel-based workflows and Power BI dashboards with an interactive, user-friendly interface that supports real-time analysis, statistical decision-making, and automated email reporting.

The application processes multiple warehouse data sources, identifies aging and statistically critical inventory, and enables department-level reporting with automated Excel exports and email delivery.

## 🎯 Business Problem

Warehouse inventory was previously analyzed using:

- Large Excel spreadsheets
- Manual filtering and reporting
- Power BI dashboards that were difficult for non-technical users to modify

This resulted in:

- High reporting time
- Risk of human error
- Limited flexibility for managers


## ✅ Solution

A Python-based analytical dashboard that:

- Accepts raw Excel exports directly from operational systems
- Automatically cleans, processes, and aggregates data
- Applies statistical thresholds to detect critical stock
- Provides interactive dashboards, filters, and downloads
- Sends automated department-specific email reports

## System Architecture

Workflow Overview:
- User uploads two Excel files (Stock Source & Fabric Stock)
- Data is validated, cleaned, and processed
- Inventory aging is calculated using date differences
- Items are categorized into time buckets
- Statistical confidence intervals identify critical inventory (95% CI)
- Results are visualized and exported
- Automated email reports are generated and sent


## Interactive Dashboard

- **KPI cards**
- **Bar charts**
- **Pivot tables**
- **Detailed item-level views**
- **CSV and ZIP export functionality**
- **Automated Email Reporting**
- **Excel reports per warehouse**
- **Gmail API integration**

---
---
## 🧠 Key Skills Demonstrated

- **Data Engineering:** Data cleaning, aggregation, pivot tables
- **Statistical Analysis:** Confidence intervals, critical threshold detection
- **Data Visualization:** Interactive dashboards, charts, KPI cards
- **Backend Logic:** Session state management, modular data pipelines
- **Automation:** Email reporting with Excel attachments
- **Product Thinking:** UX-focused design for non-technical users
