# Snowflake AI for FinOps — Cortex in Practice

A 90-minute hands-on training that teaches practitioners to understand, use, track, and govern AI spend on Snowflake using Cortex AI functions.

## What This Training Covers

Snowflake Cortex is both the tool that generates AI spend AND the platform to govern it. This training builds the skills to do both:

- **Token economics** — how tokens map to Snowflake credits across model tiers
- **Cortex AI functions** — SENTIMENT, COMPLETE, CLASSIFY_TEXT, SUMMARIZE, TRANSLATE
- **Cost-aware model selection** — choosing the cheapest model that meets your quality bar
- **Live cost tracking** — querying token and credit consumption from account_usage views
- **FinOps dashboard** — building a Streamlit in Snowflake dashboard for ongoing governance

## Prerequisites

- Snowflake account (trial or enterprise)
- ACCOUNTADMIN or SYSADMIN role access
- Web browser with Snowsight access
- Terminal access for Cortex Code CLI (optional)

## Getting Started

1. Clone this repository
2. Open `index.html` in a browser to access the training hub
3. Follow modules 01-08 sequentially

No build step or server required — all modules are static HTML.

## Modules

| # | Module | Duration | Type |
|---|--------|----------|------|
| 01 | AI Token Economy | 10 min | Concept |
| 02 | Cortex AI Capabilities | 15 min | Concept |
| 03 | Cortex Code Setup | 5 min | Setup |
| 04 | Environment Setup | 10 min | Lab |
| 05 | AI SQL Hands-On | 20 min | Lab |
| 06 | Token Usage Tracking | 12 min | Lab |
| 07 | AI Credits Transition | 5 min | Lab |
| 08 | Streamlit Dashboard | 13 min | Lab |
| 09 | Closing Note | 5 min | Wrap-up |

## File Structure

```
ai-finops-training/
├── index.html                          # Training hub with agenda and navigation
├── css/
│   └── theme.css                       # Snowflake design system
├── js/
│   └── main.js                         # Copy-to-clipboard, progress tracking
├── modules/
│   ├── 01-ai-token-economy.html        # Token economics and credit structure
│   ├── 02-cortex-ai-capabilities.html  # Cortex function catalog and cost tiers
│   ├── 03-cortex-code-setup.html       # Cortex Code CLI installation
│   ├── 04-environment-setup.html       # Database, warehouse, resource monitor, role setup
│   ├── 05-ai-sql-hands-on.html         # SQL exercises with Cortex AI functions
│   ├── 06-token-usage-tracking.html    # Cost tracking queries via account_usage
│   ├── 07-ai-credits-transition.html   # AI credits transition & governance principles
│   ├── 08-streamlit-dashboard.html     # Build and deploy a FinOps dashboard
│   └── 09-closing-note.html            # Summary and resources
├── sql/
│   ├── 04-environment-setup.sql        # Database, schema, warehouse, grants
│   ├── 05-ai-queries.sql               # Cortex AI SQL exercises
│   ├── 06-token-tracking.sql           # Usage and cost tracking queries
│   └── 07-streamlit-app.py             # Complete Streamlit dashboard code
└── README.md
```

## SQL Scripts

The `sql/` directory contains standalone scripts that can be run directly in Snowsight:

- **04-environment-setup.sql** — Creates the `cortex_lab` database, `ai_workshop` schema, `cortex_wh` warehouse, resource monitor, and `cortex_analyst` role with Cortex access
- **05-ai-queries.sql** — Five exercises covering SENTIMENT, CLASSIFY_TEXT, COMPLETE (single and multi-model comparison), and scale simulation
- **06-token-tracking.sql** — Queries for monitoring AI spend by user, function, model, and time period
- **07-streamlit-app.py** — Full Streamlit in Snowflake dashboard for ongoing AI cost governance

## Key Concepts

### Two Cost Layers
AI workloads on Snowflake have two independent cost components:
1. **Warehouse compute credits** — for running the query
2. **AI token credits** — for the Cortex AI function itself (charged per token, converted to credits)

Token credits accumulate even when the warehouse is suspended.

### Model Cost Tiers
| Tier | Examples | Relative Cost |
|------|----------|---------------|
| Budget | mistral-7b, gemma-7b | 1x |
| Standard | llama3-70b, mistral-large | 5-10x |
| Premium | claude-3-5-sonnet, llama3.1-405b | 20-50x |

### account_usage Latency
The `snowflake.account_usage` views have ~45 minutes of latency. For real-time monitoring during active sessions, use `INFORMATION_SCHEMA` alternatives (covered in Module 06).

## Facilitator Notes

### Running the Training
- Plan for 90 minutes total (5-minute buffer built in)
- Module 04 is critical — don't skip the Resource Monitor setup
- Module 05 Exercise 4 (model cost comparison) is the core FinOps lesson
- Remind participants about account_usage latency before Module 06
- Module 07 can be demo-only if time is short

### Environment Reset
To clean up between training sessions:

```sql
USE ROLE ACCOUNTADMIN;
DROP DATABASE IF EXISTS cortex_lab;
DROP WAREHOUSE IF EXISTS cortex_wh;
DROP RESOURCE MONITOR IF EXISTS cortex_lab_monitor;
DROP ROLE IF EXISTS cortex_analyst;
-- Then re-run sql/04-environment-setup.sql
```

### Common Issues

| Issue | Solution |
|-------|----------|
| Cortex functions unavailable | Check account region — not all functions are available in all regions |
| account_usage returns no data | ~45 min latency; use INFORMATION_SCHEMA for real-time data |
| Model not available | Substitute with region-available alternatives (e.g., mistral-large for claude-3-5-sonnet) |
| Resource Monitor not triggering | SUSPEND_IMMEDIATE only fires when threshold is crossed during query execution |
| Streamlit app won't run | Verify warehouse is running and user has USAGE grants on warehouse and account_usage views |

### Adding New Modules
1. Create `modules/XX-module-name.html` using an existing module as template
2. Update the navigation links in ALL existing module files
3. Update the progress bar percentage: `(module_number / total_modules) * 100`
4. Update `index.html` agenda table
5. Add corresponding SQL in `sql/` if needed

## Resources

- [Snowflake Credit Consumption Table (PDF)](https://www.snowflake.com/legal-files/CreditConsumptionTable.pdf)
- [Cortex Code CLI Getting Started](https://www.snowflake.com/en/developers/guides/getting-started-with-cortex-code-cli/)
- [FinOps Foundation](https://www.finops.org)

## License

This training material is provided for educational purposes. Feel free to adapt for internal training within your organization.
