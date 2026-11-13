# 📊 HR Employee Attrition Analysis Dashboard | Microsoft Excel

An interactive, two-tab HR analytics dashboard built entirely in Microsoft Excel to explore employee demographics, workforce distribution, organizational tenure, compensation, and employee attrition patterns.

This project transforms raw HR records into structured, standardized data and actionable visual insights through Excel formulas, PivotTables, PivotCharts, and interactive slicers. The objective is to help understand workforce composition, identify employee attrition patterns, and highlight areas that may require further investigation by HR management.

---

## 📑 Table of Contents

1. [Project Overview](#-project-overview)
2. [Business Objectives](#-business-objectives)
3. [Tools & Technologies](#-tools--technologies)
4. [Dataset Overview](#-dataset-overview)
5. [Data Cleaning & Preparation](#-data-cleaning--preparation)
6. [Dashboard Overview](#-dashboard-overview)
   - [Tab 1: Demographic & Regional Attrition Dynamics](#-tab-1-demographic--regional-attrition-dynamics)
   - [Tab 2: Temporal, Role-Based & Compensation Insights](#-tab-2-temporal-role-based--compensation-insights)
7. [Key Analytical Findings](#-key-analytical-findings)
8. [Business Recommendations](#-business-recommendations)
9. [Project Workflow](#-project-workflow)
10. [Repository Structure](#-repository-structure)
11. [Project Limitations](#-project-limitations)
12. [Conclusion](#-conclusion)

---

## 🎯 Project Overview

Employee attrition is an important challenge for organizations because workforce turnover can affect productivity, recruitment costs, employee engagement, and organizational stability.

The objective of this project is to analyze employee attrition patterns across different workforce dimensions using Microsoft Excel.

The analysis explores:

- Employee demographics and gender distribution.
- Regional workforce composition and attrition.
- Employee tenure and retention patterns.
- Age-wise attrition distribution.
- Hiring-year trends.
- Job title-wise attrition.
- Salary distribution and attrition counts.

The final deliverable is an interactive Excel dashboard consisting of two analytical tabs, supported by spreadsheet-based data cleaning, transformation, PivotTables, and visualizations.

---

## 🎯 Business Objectives

The project aims to answer the following HR-related business questions:

1. What proportion of employees have left the organization compared with those who stayed?
2. How does attrition vary across gender groups?
3. Which geographical regions show comparatively higher attrition?
4. At which stages of organizational tenure is employee turnover more prominent?
5. How does attrition differ across age groups?
6. How do attrition patterns vary across different hiring years?
7. Which job titles have comparatively higher attrition?
8. How are employees distributed across different salary brackets?
9. Which workforce segments may require further investigation to improve employee retention?

---

## 🛠️ Tools & Technologies

The entire project was developed using Microsoft Excel.

| Tool / Feature         | Application                                                     |
| ---------------------- | --------------------------------------------------------------- |
| Microsoft Excel        | Data cleaning, preparation, analysis, and dashboard development |
| Excel Formulas         | Data standardization and transformation                         |
| IF / IFERROR           | Handling blank and unmatched values                             |
| VLOOKUP                | Mapping inconsistent categorical values                         |
| TRIM / UPPER           | Removing unnecessary spaces and standardizing text              |
| SUBSTITUTE / MID / LEN | Phone number formatting and validation                          |
| PivotTables            | Data aggregation and summarization                              |
| PivotCharts            | Visual representation of analytical findings                    |
| Slicers                | Interactive dashboard filtering                                 |

---

## 📁 Dataset Overview

The project uses an HR employee dataset containing employee-level information across demographic, geographical, organizational, compensation, and employment-related attributes.

### Dataset Attributes

| Category                 | Fields                               |
| ------------------------ | ------------------------------------ |
| Employee Information     | Employee ID, Full Name               |
| Demographics             | Age, Gender                          |
| Geography                | Region                               |
| Employment               | Hire Date, Years at Company          |
| Organizational Structure | Department, Job Title                |
| Compensation             | Salary, Salary Group                 |
| Employee Experience      | Job Satisfaction, Performance Rating |
| Workforce Activity       | Monthly Hours                        |
| Employment Outcome       | Attrition                            |
| Additional Information   | Payment Method, Phone Number         |

The raw dataset contained inconsistent categorical values, different formatting conventions, and missing entries that required preparation before analysis.

---

## 🧹 Data Cleaning & Preparation

Data cleaning and standardization were performed entirely in Microsoft Excel.

The primary objective was to improve data consistency, standardize categorical values, validate phone number formats, and prepare the dataset for accurate analysis.

Helper columns were created using Excel formulas, and the resulting values were reviewed before being used in the final analytical dataset.

### 1. Region Standardization (Q → R)

**Problem:**

Regional information contained inconsistent abbreviations, spelling variations, and capitalization.

**Solution:**

Standardized regional values into five consistent categories.

| Original Values               | Standardized Value |
| ----------------------------- | ------------------ |
| EUR, EUROPE                   | EUROPE             |
| NA, N. AMERICA, NORTH AMERICA | NORTH AMERICA      |
| APAC, ASIA PACIFIC            | ASIA PACIFIC       |
| ME, MID EAST, MIDDLE EAST     | MIDDLE EAST        |
| LATAM, LATIN AMERICA          | LATIN AMERICA      |

**Excel Formula (R2):**

```excel
=IF(TRIM(Q2)="","",IFERROR(VLOOKUP(UPPER(TRIM(Q2)),{"EUR","EUROPE";"EUROPE","EUROPE";"NA","NORTH AMERICA";"N. AMERICA","NORTH AMERICA";"NORTH AMERICA","NORTH AMERICA";"APAC","ASIA PACIFIC";"ASIA PACIFIC","ASIA PACIFIC";"ME","MIDDLE EAST";"MID EAST","MIDDLE EAST";"MIDDLE EAST","MIDDLE EAST";"LATAM","LATIN AMERICA";"LATIN AMERICA","LATIN AMERICA"},2,FALSE),"CHECK"))
```

**Cleaning logic:**

- TRIM removes unnecessary leading and trailing spaces.
- UPPER standardizes capitalization.
- VLOOKUP maps recognized variations to their standard values.
- IFERROR identifies unmatched values using `CHECK`.
- Blank source cells remain blank.

### 2. Payment Method Standardization (S → T)

**Problem:**

Payment methods contained inconsistent capitalization and abbreviated entries.

**Solution:**

Standardized payment methods into four categories.

| Original Values     | Standardized Value |
| ------------------- | ------------------ |
| DIRECT DEPOSIT, DD  | DIRECT DEPOSIT     |
| CASH                | CASH               |
| WIRE TRANSFER, WIRE | WIRE TRANSFER      |
| CHECK, CHK          | CHECK              |

**Excel Formula (T2):**

```excel
=IF(TRIM(S2)="","",IFERROR(VLOOKUP(UPPER(TRIM(S2)),{"DIRECT DEPOSIT","DIRECT DEPOSIT";"DD","DIRECT DEPOSIT";"CASH","CASH";"WIRE TRANSFER","WIRE TRANSFER";"WIRE","WIRE TRANSFER";"CHECK","CHECK";"CHK","CHECK"},2,FALSE),"CHECK THIS"))
```

**Cleaning logic:**

- TRIM removes unnecessary spaces.
- UPPER converts entries into uppercase.
- VLOOKUP maps abbreviations to standardized payment methods.
- IFERROR returns `CHECK THIS` for unmatched values.

### 3. Phone Number Standardization (U → V)

**Problem:**

Phone numbers were stored in multiple formats, including dots, spaces, hyphens, and parentheses.

**Solution:**

Standardized valid 10-digit phone numbers into a consistent format.

**Expected Output Format:** `(608) 834-6257`

| Original Value | Standardized Value |
| -------------- | ------------------ |
| 608.834.6257   | (608) 834-6257     |
| 774-262-5380   | (774) 262-5380     |
| 334 259 8799   | (334) 259-8799     |
| (747) 238-5524 | (747) 238-5524     |

**Excel Formula (V2):**

```excel
=IF(TRIM(U2)="","",IF(LEN(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(U2,"(",""),")","")," ",""),"-",""),".",""))=10,"("&MID(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(U2,"(",""),")","")," ",""),"-",""),".",""),1,3)&") "&MID(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(U2,"(",""),")","")," ",""),"-",""),".",""),4,3)&"-"&MID(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(U2,"(",""),")","")," ",""),"-",""),".",""),7,4),"CHECK THIS"))
```

**Cleaning logic:**

- SUBSTITUTE removes common formatting characters.
- LEN checks whether the number contains exactly 10 digits after cleaning.
- MID extracts the area code, middle digits, and final four digits.
- The formula reconstructs the phone number into a consistent format.
- Invalid entries return `CHECK THIS`.

### 4. Attrition Standardization (O → P)

**Problem:**

Employee attrition status was represented using different variations, including YES, NO, Y, N, LEFT, and STAYED.

**Solution:**

Standardized attrition values into two categories.

| Original Values | Standardized Value |
| --------------- | ------------------ |
| YES, Y, LEFT    | YES                |
| NO, N, STAYED   | NO                 |

**Excel Formula (P2):**

```excel
=IF(TRIM(O2)="","",IFERROR(VLOOKUP(UPPER(TRIM(O2)),{"YES","YES";"Y","YES";"LEFT","YES";"NO","NO";"N","NO";"STAYED","NO"},2,FALSE),"CHECK THIS"))
```

**Interpretation:**

- YES = Employee left the organization.
- NO = Employee remained with the organization.

Unrecognized values return `CHECK THIS` for further review.

### Data Cleaning Principles

The following practices were applied during data preparation:

- Standardized inconsistent categorical values.
- Removed unnecessary leading and trailing spaces.
- Applied consistent capitalization.
- Preserved blank source values as blank outputs.
- Flagged unmatched values for manual review.
- Used helper columns to validate transformations.
- Prepared standardized values for PivotTable analysis.

**Final Cleaning Step:**

1. Filled the formulas down across the dataset.
2. Reviewed unmatched values.
3. Checked the standardized output.
4. Used Copy and Paste Special → Values to finalize cleaned data where appropriate.
5. Used the prepared data for PivotTables and dashboard development.

---

## 📊 Dashboard Overview

The final Excel dashboard consists of two interactive tabs.

Both tabs follow a consistent dark-mode design with high-contrast text, orange highlights, and steel-blue chart elements.

### 📍 Tab 1: Demographic & Regional Attrition Dynamics

This dashboard focuses on employee demographics, geographical distribution, and organizational tenure.

### Interactive Slicers

| Slicer           | Available Categories                                                 |
| ---------------- | -------------------------------------------------------------------- |
| Gender           | Male, Female, N/A                                                    |
| Region           | Asia Pacific, Europe, Latin America, Middle East, North America, N/A |
| Years at Company | 0–5, 6–10, 11–15, 16–20, 21–24 years, N/A                            |

These slicers allow users to dynamically filter the dashboard and explore specific workforce segments.

### Visualizations

**1. Overall Employee Attrition – Pie Chart**

**Purpose:** Displays the overall proportion of employees who stayed and those who left the organization.

| Attrition Status | Percentage |
| ---------------- | ---------: |
| NO – Retained    |     50.69% |
| YES – Attrited   |     49.31% |

**Key Observation:** The displayed distribution indicates a nearly equal split between retained and attrited employees.

**2. Gender-wise Attrition – Clustered Column Chart**

**Purpose:** Compares retention and attrition distribution across gender groups.

| Gender | NO – Retained | YES – Attrited |
| ------ | ------------: | -------------: |
| Female |        50.64% |         49.36% |
| Male   |        50.77% |         49.23% |

**Key Observation:** The displayed proportions are very similar across both gender groups.

**3. Region-wise Attrition – Clustered Column Chart**

**Purpose:** Evaluates employee attrition patterns across five geographical regions.

**Key Observation:**

- Latin America records a displayed YES proportion of 48.33%.
- Regional comparisons provide visibility into geographical differences in workforce retention.

**4. Attrition by Years at Company – Clustered Column Chart**

**Purpose:** Examines how employee attrition patterns vary across different organizational tenure brackets.

**Key Observation:**

- The 11–15 years tenure group records a displayed YES proportion of 48.14%.
- The 6–10 years group is another important segment for retention analysis.
- The visualization helps identify tenure groups that may require further investigation.

---

### 📍 Tab 2: Temporal, Role-Based & Compensation Insights

The second dashboard explores employee attrition through age groups, hiring years, job titles, and salary brackets.

### Interactive Slicers

| Slicer    | Available Categories                                                                                |
| --------- | --------------------------------------------------------------------------------------------------- |
| Hire Year | 2018–2023, N/A                                                                                      |
| Job Title | Analyst, Associate, Consultant, Director, Intern, Lead, Manager, Senior Analyst, Senior Manager, VP |

These slicers allow users to investigate specific recruitment cohorts and organizational roles.

### Visualizations

**1. Age-wise Attrition – Column Chart**

**Purpose:** Identifies how attrition distribution varies across employee age groups.

| Age Group | Attrition Percentage |
| --------- | -------------------: |
| 21–30     |               23.96% |
| 31–40     |               25.24% |
| 41–50     |               24.76% |
| 51–61     |               26.05% |

**Key Observation:**

- The 51–61 age group records the highest displayed percentage at 26.05%.
- Differences across age groups are relatively small.

**2. Year-wise Attrition – Clustered Column Chart**

**Purpose:** Tracks historical attrition patterns across years from 2010 to 2023.

**Key Observations:**

- 2011 records a displayed YES proportion of 45.52% and NO proportion of 51.59%.
- 2023 records a displayed YES proportion of 53.61%.
- The visualization allows comparison of attrition patterns across different hiring-year cohorts.

**3. Job Title-wise Attrition – Clustered Column Chart**

**Purpose:** Evaluates employee retention and attrition across organizational designations.

**Key Observations:**

- Senior Analysts record a displayed YES proportion of 52.25%.
- Managers record a displayed YES proportion of 50.99%.
- Role-based comparisons help identify positions that may require further investigation into retention challenges.

**4. Salary-wise Attrition – Clustered Column Chart**

**Purpose:** Examines employee distribution and attrition counts across salary brackets.

Salary categories include:

- 25,000–50,000
- 50,001–75,000
- 75,001–100,000
- 100,001–125,000
- Higher salary brackets, including above 168,000.

**Key Observations:**

- The 25,000–50,000 and 50,001–75,000 brackets contain the highest employee counts.
- These brackets also account for substantial attrition volumes.
- The visualization helps compare workforce distribution and attrition counts across compensation categories.

---

## 📈 Key Analytical Findings

The dashboards highlight several workforce patterns worth investigating.

| Analytical Area   | Key Finding                                                                                 |
| ----------------- | ------------------------------------------------------------------------------------------- |
| Overall Attrition | 49.31% displayed YES proportion                                                             |
| Gender            | Similar attrition distribution across male and female employees                             |
| Region            | Latin America records a 48.33% YES proportion                                               |
| Tenure            | 11–15 years records a 48.14% YES proportion                                                 |
| Age               | 51–61 age group records the highest displayed age-wise percentage at 26.05%                 |
| Hiring Year       | 2023 records a displayed YES proportion of 53.61%                                           |
| Job Title         | Senior Analysts record a 52.25% YES proportion                                              |
| Salary            | Lower salary brackets contain the highest employee counts and substantial attrition volumes |

These findings are descriptive observations from the available dataset. They should not be interpreted as proof of the causes of employee turnover.

---

## 💡 Business Recommendations

Based on the observed patterns, the following areas could be investigated by HR management:

**1. Employee Retention**

Investigate the factors contributing to the high overall attrition proportion and identify opportunities to improve employee engagement.

**2. Regional Workforce Analysis**

Examine regional differences in working conditions, employee satisfaction, compensation, and management practices.

**3. Tenure-based Retention**

Evaluate employee satisfaction, career progression, and promotion opportunities across different organizational tenure groups.

**4. Role-based Retention**

Investigate retention challenges among Senior Analysts and Managers through employee feedback and exit interviews.

**5. Compensation Analysis**

Review salary structures and compensation competitiveness, particularly in lower salary brackets.

**6. Recruitment Cohort Analysis**

Compare hiring-year cohorts to understand how employee retention patterns differ across recruitment periods.

These recommendations are potential areas for further investigation rather than confirmed causes of attrition.

---

## 🔄 Project Workflow

The project followed a structured Excel-based analytical workflow.

| Step | Activity              | Description                                                     |
| ---- | --------------------- | --------------------------------------------------------------- |
| 1    | Data Understanding    | Reviewed the raw HR dataset and identified relevant fields      |
| 2    | Data Cleaning         | Standardized inconsistent values using Excel formulas           |
| 3    | Data Transformation   | Created helper columns and standardized categorical information |
| 4    | Data Validation       | Reviewed unmatched values and checked cleaned outputs           |
| 5    | Data Analysis         | Used PivotTables to summarize employee information              |
| 6    | Dashboard Development | Created PivotCharts and interactive slicers                     |
| 7    | Insight Generation    | Identified workforce and attrition patterns                     |
| 8    | Documentation         | Documented the methodology, findings, and recommendations       |

---

## 📂 Repository Structure

```text
HR-Employee-Attrition-Analysis/
│
├── README.md
│
├── Dataset/
│   └── HR_Employee_Attrition.xlsx
│
├── Excel/
│   └── HR_Employee_Attrition_Dashboard.xlsx
│
└── Dashboard_Screenshots/
    ├── Dashboard_Tab_1.png
    └── Dashboard_Tab_2.png
```

*The structure above is a suggested repository organization. Adjust filenames to match the files actually uploaded.*

---

## ⚠️ Project Limitations

- The analysis is limited to the attributes available in the dataset.
- Missing and unmatched values may affect certain category-level comparisons.
- The dashboard identifies patterns and relationships but does not establish causation.
- Salary-based employee counts should not be confused with salary-specific attrition rates.
- Additional information, such as exit interview feedback, employee engagement surveys, and compensation benchmarks, would be required to investigate the underlying causes of attrition.
- The findings reflect the dataset used for this project and should not automatically be generalized to other organizations.

---

## 🏁 Conclusion

The HR Employee Attrition Analysis Dashboard demonstrates how Microsoft Excel can be used to transform raw workforce data into structured information and interactive visual reports.

Through formula-based data cleaning, standardization, PivotTables, PivotCharts, and slicers, this project explores employee attrition across demographic, geographical, tenure, recruitment, job title, and compensation dimensions.

The resulting dashboards provide a consolidated view of workforce patterns and highlight areas where further HR investigation may support better employee retention decisions.

---

**Project Category:** Data Analytics | HR Analytics | Workforce Analytics

**Technology Used:** Microsoft Excel

**Core Skills Demonstrated:** Data Cleaning | Data Standardization | Excel Formulas | PivotTables | PivotCharts | Interactive Dashboards | Data Analysis | Business Insights
