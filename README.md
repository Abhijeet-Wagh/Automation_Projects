# Automation of Commissions Review Report

## Project Overview
This project automates the **Commissions Review Report**, a critical monthly process designed to identify high-achieving Sales Representatives eligible for commission review. By automating data filtering, validation, and report generation, this solution streamlines the preparation process, ensuring timely and accurate information for the Operations Leads who analyze each case.

## Project Goals
- **Efficiency**: Automate the manual, repetitive steps to significantly reduce the preparation time.
- **Accuracy**: Eliminate potential human errors in data entry and formatting.
- **Standardization**: Ensure each report follows a consistent format, simplifying analysis.
- **Scalability**: Handle multiple cases across different regions effortlessly.

## Automation Workflow

### Key Steps
1. **Data Processing**  
   - Import the **Stack Rank Report** as input, then filter and extract relevant data fields for cases that meet performance thresholds:
     - **Commissions Due** > $50K USD
     - **Order-to-Cash (OTC)** > 250%
     - **Quota Attainment** > 150%
  
2. **Template Preparation**  
   - Download the assignment report and prepare a template containing:
     - Bookings data for each Sales Rep
     - A formatted, blank sheet to document case details

3. **Data Population with VBA Macro**  
   - Run the VBA Macro to:
     - Populate data for each Sales Rep into the template, filling in key metrics such as bookings, quota, pipeline, forecast, attainment rates, and deal-level details.
     - Create a tab for each Sales Rep’s case, renaming it with their name for easy identification.

4. **Output Generation**  
   - Generate an Excel file with individual tabs for each Sales Rep, organized by region.
   - Distribute the regional files to respective Ops Leads for review, including a structured questionnaire for collecting feedback and recommendations.

5. **Consolidation and Presentation**  
   - Once Ops Leads complete their research, consolidate responses into a final report.
   - Present the updated report during the **High Achievers Commissions Review** call.

## Technical Stack
- **Python**: Data filtering and preprocessing of the Stack Rank Report.
- **Excel VBA**: Automated data population, formatting, and generation of individual tabs for each Sales Rep’s case.

## Project Benefits

- **Time-Saving**: Reduces manual preparation time, ensuring reports are ready before the tight review deadlines.
- **Enhanced Accuracy**: Automates data entry, minimizing errors and increasing data reliability.
- **Improved Decision-Making**: Provides Ops Leads with structured, detailed information, enabling better analysis and informed recommendations.
- **Repeatability and Scalability**: Supports a scalable solution for handling multiple cases, regions, and large datasets.

## How to Use
1. **Run Python Script**: Filter the Stack Rank Report to obtain cases meeting the set criteria.
2. **Download and Prepare Assignment Report**: Ensure the template is complete and ready for macro use.
3. **Run VBA Macro**: Populate cases in the required format and generate the regional Excel files.
4. **Distribute and Consolidate**: Send files to Ops Leads, then collect and consolidate inputs into a final report for review.

## Future Enhancements
- Automate the regional distribution step.
- Integrate the tool with a dashboard for real-time tracking and report updates.

---

This README provides a structured overview of the project, including its purpose, functionality, technical stack, and benefits, making it easy for others to understand the scope and value of your automation work.
