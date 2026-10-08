# LPS-Beacon

Description: 
The Resource Planning Forecast Chatbot is an AI-powered assistant that enables users to upload project, pipeline, workforce, and financial planning data to generate headcount forecasts, analyze resource demand and supply, perform resource matching, and produce monthly forecast reports.

The chatbot serves as the primary user interface, allowing users to interact using natural language. Behind the scenes, AWS services process uploaded files, transform data, execute forecasting models, and generate reports through the Resource Planning Forecast Agent.



Chatbot Interface for Resource Planning 
User
  │
  ▼
Chatbot Interface
  │
  ▼
Upload Files
(Project Pipeline,
Resource Supply,
Forecast Inputs)
  │
  ▼
Amazon S3
  │
  ▼
SQS Notification Queue
  │
  ▼
Upload Processing Lambda
  │
  ▼
AWS Glue ETL(Lambda)
(Read, Clean, Transform)
  │
  ▼
Transformed Data Store
(S3) - Processed Buckets
  │
  ▼
Forecast Agent Lambda / BedRock Model(LLM)
  │
  ├── Demand Forecast
  ├── Headcount Forecast
  ├── Resource Matching
  └── Capacity Analysis - Capacity Planning
  │
  ▼
Athena Query Engine
  │
  ▼
Forecast Results(Bucket csv.)
  │
  ▼
Report Generation Lambda
  │
  ▼
Monthly Forecast Report(Excel/CSV)
(PDF / Excel / CSV)
  │
  ▼
S3 Output Bucket
  │
  ▼
SNS Notification
  │
  ▼
Chatbot Returns User Prompt Queries 
`
