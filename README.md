Power BI Metadata Extractor – Desktop App & PowerShell Automation

Built an end-to-end Power BI Admin API solution to extract tenant-wide architecture metadata, including workspaces, reports, datasets, Power Query M expressions, DAX measures, model relationships, RLS roles, and artifact lineage.
Developed a Windows desktop application (Python, CustomTkinter, PyInstaller) with progress tracking, activity logging, prerequisite checks, and one-click Excel export for non-technical users.
Delivered a pure PowerShell alternative for enterprise clients that restrict .exe execution, using only official Microsoft modules (MicrosoftPowerBIMgmt, Admin Scanner API).
Implemented robust authentication via Microsoft sign-in, silent/secure token handling, and Admin API validation (role + tenant settings).
Designed large-tenant scanning with batched workspace processing, fault isolation, and graceful fallbacks so one failed dataset does not stop the full extraction.
Parsed Power Query M code to detect connection types (SQL, Teradata, BigQuery, Fabric, Excel, etc.), connection strings, and implicit joins (Table.NestedJoin / Table.Join).
Generated structured multi-sheet Excel reports (13 sheets) plus JSON backups for architecture documentation, audit, and migration planning.
Produced client-facing setup guides, prerequisite documentation, and troubleshooting steps for production rollout
