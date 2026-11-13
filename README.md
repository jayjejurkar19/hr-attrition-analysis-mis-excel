# 📊 HR Employee Attrition Analysis Dashboard

An interactive, two-page HR analytics dashboard developed to analyze employee demographics, workforce distribution, organizational tenure, compensation, and employee attrition patterns.

This project transforms raw HR records into structured, standardized data and meaningful visual insights using **Microsoft Excel, MySQL, and data visualization techniques**. It explores workforce trends, identifies employee attrition patterns across different organizational dimensions, and provides management with a consolidated view of workforce composition and retention.

---

## 📌 Table of Contents

1. [Project Overview](#-project-overview)
2. [Business Objectives](#-business-objectives)
3. [Tools & Technologies](#-tools--technologies)
4. [Dataset Overview](#-dataset-overview)
5. [Data Cleaning & Preparation](#-data-cleaning--preparation)
6. [Dashboard Overview](#-dashboard-overview)
   - [Tab 1: Demographic & Regional Analysis](#-tab-1-demographic--regional-attrition-dynamics)
   - [Tab 2: Temporal, Role-Based & Compensation Analysis](#-tab-2-temporal-role-based--compensation-insights)
7. [Key Analytical Findings](#-key-analytical-findings)
8. [SQL Validation](#-sql-validation)
9. [Business Recommendations](#-business-recommendations)
10. [Project Workflow](#-project-workflow)
11. [Repository Structure](#-repository-structure)
12. [Limitations](#-limitations)

---

## 🎯 Project Overview

Employee attrition is an important workforce management challenge. Understanding why employees leave, which employee segments experience higher turnover, and how workforce characteristics relate to retention can help organizations make better HR decisions.

The objective of this project is to explore employee attrition patterns across multiple dimensions, including:

- Gender and regional distribution.
- Employee age groups.
- Years of service within the organization.
- Hiring trends and recruitment years.
- Job titles and organizational hierarchy.
- Salary distribution and compensation brackets.

The final solution consists of two interactive dashboard tabs, supported by spreadsheet-based data cleaning and SQL-based analytical validation.

## 🎯 Business Objectives

The project focuses on answering the following business questions:

- What proportion of employees have left the organization?
- How does attrition differ across gender groups?
- Which geographical regions experience higher attrition?
- At which organizational tenure stages is employee turnover more prominent?
- How does attrition vary across different age groups?
- How have attrition patterns changed across hiring years?
- Which job titles experience comparatively higher attrition?
- How are employees distributed across different salary brackets?

---

## 🛠️ Tools & Technologies

| Tool | Purpose |
|---|---|
| Microsoft Excel | Data cleaning, standardization, and transformation |
| MySQL | Data validation, aggregation, and analytical queries |
| Excel PivotTables | Data summarization and dashboard data preparation |
| Excel Charts | Interactive data visualization |
| Excel Slicers | Dynamic filtering and dashboard interaction |
| GitHub | Project documentation and version control |

---

## 📁 Dataset Overview

The project uses an HR employee dataset containing employee-level information across demographic, organizational, compensation, and employment-related attributes.

The dataset includes fields such as:

| Category | Attributes |
|---|---|
| Employee Information | Employee ID, Full Name |
| Demographics | Age, Gender |
| Geography | Region |
| Employment | Hire Date, Years at Company |
| Organizational Structure | Department, Job Title |
| Compensation | Salary, Salary Group |
| Employee Experience | Job Satisfaction, Performance Rating |
| Workforce Activity | Monthly Hours |
| Employment Outcome | Attrition |

The dataset was prepared for analysis by standardizing inconsistent values, handling missing entries, and creating structured analytical categories.

---

## 🧹 Data Cleaning & Preparation

Data preparation was performed in Microsoft Excel before building the analytical dashboards.

The objective was to improve data consistency, reduce duplicate category representations, and ensure that the dataset could be accurately summarized using PivotTables and SQL queries.

### 1. Region Standardization

**Problem:** Regional information contained inconsistent abbreviations, naming conventions, and capitalization.

**Solution:** Standardized regional values into five consistent categories.

| Original Values | Standardized Value |
|---|---|
| EUR, EUROPE | EUROPE |
| NA, N. AMERICA, NORTH AMERICA | NORTH AMERICA |
| APAC, ASIA PACIFIC | ASIA PACIFIC |
| ME, MID EAST, MIDDLE EAST | MIDDLE EAST |
| LATAM, LATIN AMERICA | LATIN AMERICA |

**Excel Formula (R2):**

```excel
=IF(TRIM(Q2)="","",IFERROR(VLOOKUP(UPPER(TRIM(Q2)),{"EUR","EUROPE";"EUROPE","EUROPE";"NA","NORTH AMERICA";"N. AMERICA","NORTH AMERICA";"NORTH AMERICA","NORTH AMERICA";"APAC","ASIA PACIFIC";"ASIA PACIFIC","ASIA PACIFIC";"ME","MIDDLE EAST";"MID EAST","MIDDLE EAST";"MIDDLE EAST","MIDDLE EAST";"LATAM","LATIN AMERICA";"LATIN AMERICA","LATIN AMERICA"},2,FALSE),"CHECK"))
```

Unrecognized values return `CHECK` for further investigation.

### 2. Payment Method Standardization

**Problem:** Payment methods contained inconsistent capitalization and abbreviated entries.

**Solution:** Standardized payment methods into four categories.

| Original Values | Standardized Value |
|---|---|
| DIRECT DEPOSIT, DD | DIRECT DEPOSIT |
| CASH | CASH |
| WIRE TRANSFER, WIRE | WIRE TRANSFER |
| CHECK, CHK | CHECK |

**Excel Formula (T2):**

```excel
=IF(TRIM(S2)="","",IFERROR(VLOOKUP(UPPER(TRIM(S2)),{"DIRECT DEPOSIT","DIRECT DEPOSIT";"DD","DIRECT DEPOSIT";"CASH","CASH";"WIRE TRANSFER","WIRE TRANSFER";"WIRE","WIRE TRANSFER";"CHECK","CHECK";"CHK","CHECK"},2,FALSE),"CHECK THIS"))
```

Unrecognized entries return `CHECK THIS`.

### 3. Phone Number Standardization

**Problem:** Phone numbers were stored in different formats, including dots, spaces, hyphens, and parentheses.

**Solution:** Standardized valid 10-digit phone numbers into a consistent format.

**Expected Format:** `(608) 834-6257`

| Original Value | Standardized Value |
|---|---|
| 608.834.6257 | (608) 834-6257 |
| 774-262-5380 | (774) 262-5380 |
| 334 259 8799 | (334) 259-8799 |
| (747) 238-5524 | (747) 238-5524 |

**Excel Formula (V2):**

```excel
=IF(TRIM(U2)="","",IF(LEN(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(U2,"(",""),")","")," ",""),"-",""),".",""))=10,"("&MID(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(U2,"(",""),")","")," ",""),"-",""),".",""),1,3)&") "&MID(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(U2,"(",""),")","")," ",""),"-",""),".",""),4,3)&"-"&MID(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(U2,"(",""),")","")," ",""),"-",""),".",""),7,4),"CHECK THIS"))
```

The formula removes common formatting characters and reconstructs the phone number if exactly 10 digits are present.

Invalid entries return `CHECK THIS`.

### 4. Attrition Standardization

**Problem:** Employee attrition status was represented using multiple variations, including Yes, No, Y, N, Left, and Stayed.

**Solution:** Standardized all entries into two categories.

| Original Values | Standardized Value |
|---|---|
| YES, Y, LEFT | YES |
| NO, N, STAYED | NO |

**Excel Formula (P2):**

```excel
=IF(TRIM(O2)="","",IFERROR(VLOOKUP(UPPER(TRIM(O2)),{"YES","YES";"Y","YES";"LEFT","YES";"NO","NO";"N","NO";"STAYED","NO"},2,FALSE),"CHECK THIS"))
```

**Interpretation:**
- YES = Employee left the organization.
- NO = Employee remained with the organization.

Unrecognized entries return `CHECK THIS`.

### Data Cleaning Principles

The following principles were applied throughout the cleaning process:

- Standardized inconsistent categorical values.
- Removed unnecessary leading and trailing spaces.
- Used consistent capitalization.
- Preserved blank source values as blank outputs.
- Flagged unmatched values for manual review.
- Created helper columns to validate transformations before finalizing the cleaned dataset.

After validation, the cleaned values were used for subsequent analysis.

---

## 📊 Dashboard Overview

The final dashboard is divided into two interactive tabs, each focusing on a different aspect of employee attrition.

Both tabs use a consistent dark-themed visual design with contrasting text, orange highlights, and steel-blue chart elements.

### 📍 Tab 1: Demographic & Regional Attrition Dynamics

This dashboard examines employee attrition across demographic characteristics, geographical regions, and organizational tenure.

**Dashboard Components**

#### Interactive Slicers

| Slicer | Available Categories |
|---|---|
| Gender | Male, N/A |
| Region | Asia Pacific, Europe, Latin America, Middle East, North America, N/A |
| Years at Company | 0–5, 6–10, 11–15, 16–20, 21–24 years, N/A |

These slicers allow users to filter the dashboard dynamically and explore different workforce segments.

#### Visualizations

**1. Overall Employee Attrition (Pie Chart)**

Purpose: Provides an overview of the proportion of employees who stayed and those who left.

| Attrition Status | Percentage |
|---|---:|
| NO – Retained | 50.69% |
| YES – Attrited | 49.31% |

**2. Gender-wise Attrition (Clustered Column Chart)**

Purpose: Compares the distribution of retained and attrited employees across gender groups.

| Gender | NO – Retained | YES – Attrited |
|---|---:|---:|
| Female | 50.64% | 49.36% |
| Male | 50.77% | 49.23% |

**3. Region-wise Attrition (Clustered Column Chart)**

Purpose: Examines employee attrition patterns across five geographical regions.

Key observation:
- Latin America records a YES proportion of 48.33%.
- Regional comparisons help identify differences in employee retention patterns across geographical territories.

**4. Attrition by Years at Company (Clustered Column Chart)**

Purpose: Examines how employee attrition patterns vary across organizational tenure brackets.

Key observations:
- The 11–15 years tenure group records a YES proportion of 48.14%.
- The 6–10 years group is another important tenure segment to examine.
- The analysis provides visibility into employee retention patterns across different stages of organizational service.

---

### 📍 Tab 2: Temporal, Role-Based & Compensation Insights

The second dashboard focuses on age distribution, recruitment timelines, job titles, and salary brackets to examine additional dimensions of employee attrition.

#### Interactive Slicers

| Slicer | Available Categories |
|---|---|
| Hire Year | 2018–2023, N/A |
| Job Title | Analyst, Associate, Consultant, Director, Intern, Lead, Manager, Senior Analyst, Senior Manager, VP |

These filters allow users to isolate specific recruitment cohorts and organizational roles.

#### Visualizations

**1. Age-wise Attrition (Column Chart)**

Purpose: Evaluates employee attrition patterns across different age groups.

| Age Group | Attrition Percentage |
|---|---:|
| 21–30 | 23.96% |
| 31–40 | 25.24% |
| 41–50 | 24.76% |
| 51–61 | 26.05% |

Key observation:
- The 51–61 age group records the highest displayed attrition percentage at 26.05%.
- The differences across age groups are relatively small.

**2. Year-wise Attrition (Clustered Column Chart)**

Purpose: Examines historical attrition patterns across recruitment years from 2010 to 2023.

Key observations:
- The chart highlights changes in attrition distribution across different years.
- 2011 records a displayed YES proportion of 45.52% and NO proportion of 51.59%.
- 2023 records a displayed YES proportion of 53.61%.

This visualization helps compare workforce outcomes across different hiring-year cohorts.

**3. Job Title-wise Attrition (Clustered Column Chart)**

Purpose: Compares employee retention and attrition patterns across organizational designations.

Key observations:
- Senior Analysts record a displayed YES proportion of 52.25%.
- Managers record a displayed YES proportion of 50.99%.
- Role-based comparisons help identify positions that may require further investigation into retention challenges.

**4. Salary-wise Attrition (Clustered Column Chart)**

Purpose: Examines employee distribution and attrition counts across different salary brackets.

Salary categories include:

- 25,000–50,000
- 50,001–75,000
- 75,001–100,000
- 100,001–125,000
- Higher salary brackets, including above 168,000.

Key observations:
- The 25,000–50,000 and 50,001–75,000 salary brackets contain the highest employee counts.
- These brackets also account for substantial attrition volumes.

This visualization helps examine how workforce distribution and attrition counts vary across compensation levels.

---

## 📈 Key Analytical Findings

The dashboard provides several observations worth investigating further:

| Analytical Area | Key Finding |
|---|---|
| Overall Attrition | 49.31% displayed attrition proportion |
| Gender | Similar attrition distribution across male and female employees |
| Region | Latin America records a 48.33% YES proportion |
| Tenure | The 11–15 years group records a 48.14% YES proportion |
| Age | The 51–61 age group has the highest displayed age-wise attrition percentage at 26.05% |
| Recruitment Year | The 2023 cohort records a 53.61% YES proportion |
| Job Title | Senior Analysts record a 52.25% YES proportion |
| Salary | Lower salary brackets contain the highest employee counts and substantial attrition volumes |

These findings are descriptive observations from the available dataset. They identify areas for further investigation rather than establishing the underlying causes of employee turnover.

---

## 🗄️ SQL Validation

MySQL was used to validate the prepared dataset and verify analytical aggregations.

The validation process focused on:

- Checking total employee records.
- Verifying YES and NO attrition counts.
- Comparing category-level employee distributions.
- Validating percentage calculations.
- Checking regional and demographic aggregations.
- Cross-checking analytical results against Excel PivotTables.

This additional validation helps improve confidence in the consistency of the dashboard calculations.

---

## 💡 Business Recommendations

Based on the observed patterns, the following areas could be considered for further HR investigation:

1. **Employee Retention:** Investigate the factors contributing to the high overall attrition proportion.
2. **Regional Workforce Analysis:** Examine regional differences in compensation, working conditions, and employee engagement.
3. **Tenure-based Retention:** Evaluate employee satisfaction and career progression opportunities across different tenure groups.
4. **Role-based Retention:** Investigate retention challenges among Senior Analysts and Managers.
5. **Compensation Analysis:** Review salary structures and employee satisfaction across lower compensation brackets.
6. **Recruitment Cohort Analysis:** Compare newer and older hiring cohorts to understand differences in workforce outcomes.

These recommendations should be supported by additional information, such as exit interview data, compensation benchmarks, employee engagement surveys, and workload analysis.

---

## 🔄 Project Workflow

The project followed a structured analytical workflow:

1. **Data Understanding:** Reviewed the raw HR dataset and identified relevant analytical fields.
2. **Data Cleaning:** Standardized inconsistent categorical values and phone number formats using Excel formulas.
3. **Data Validation:** Reviewed unmatched entries and checked cleaned values.
4. **Exploratory Analysis:** Used MySQL and Excel PivotTables to examine workforce distributions and attrition patterns.
5. **Dashboard Development:** Created two interactive dashboard tabs with charts and slicers.
6. **Analytical Interpretation:** Identified key workforce patterns and translated them into potential business recommendations.
7. **Documentation:** Organized the cleaning methodology, dashboard descriptions, and analytical findings for project reproducibility.

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
├── SQL/
│   └── HR_Attrition_Analysis.sql
│
└── Dashboard_Screenshots/
    ├── Dashboard_Tab_1.png
    └── Dashboard_Tab_2.png
```

*Repository structure is illustrative. File and folder names should match the actual uploaded project files.*

---

## ⚠️ Limitations

- The analysis is based on the available dataset and its recorded attributes.
- Missing and unmatched values may limit certain category-level comparisons.
- The dashboard identifies associations and distribution patterns but does not establish causation.
- Salary-based employee counts should not be interpreted as salary-specific attrition rates.
- Additional HR information would be required to identify the underlying reasons for employee departures.

---

## 👨‍💻 Project Purpose

This project demonstrates practical application of data analytics concepts, including spreadsheet-based data cleaning, data standardization, SQL validation, exploratory analysis, PivotTable reporting, and interactive dashboard development.

It highlights how raw employee records can be transformed into structured information to support workforce analysis and data-driven HR decision-making.

**Project Category:** Data Analytics | HR Analytics | Workforce Analytics

**Tools:** Microsoft Excel | MySQL | GitHub

---
