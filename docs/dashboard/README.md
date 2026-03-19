# Antibiotic Utilization Dashboard

**Work in progress:** This dashboard is still under development. The goal is to make Spectrum Score visualization and antibiotic utilization data more accessible to other Infectious Diseases pharmacists and antimicrobial stewardship programs that may find running R to be a barrier.

## Overview

This is an interactive HTML dashboard for visualizing Spectrum Scores and antibiotic utilization. It is designed to run entirely in a web browser, with no data sent or stored externally.

**Origin:** This dashboard was ported using AI from existing R code. Some modifications have been added, such as using **DASC/DOT instead of DASC/encounter** and additional customizability for each client.

## HIPAA-Safe Instructions

- Only upload **de-identified patient data**.  
- The dashboard runs entirely on the client side and does **not store or transmit any patient information**.  
- Use this tool responsibly within your stewardship program or educational setting.

## Usage

1. Open `index.html` in a web browser **or access it directly online**:  
   [View the Dashboard](https://shemtsai.github.io/semi-automated-report-cards/dashboard/)  
2. View or populate the table manually, or integrate client-side scripts for automated data visualization (if added).  
3. All CSS and JS are contained within the file or referenced via CDN; no server setup is required.

## Goal

- Make antibiotic utilization and Spectrum Score data **more accessible** for stewardship teams.  
- Provide a starting point for other programs to adapt and build upon.  

---

**Note:** This dashboard is still a work in progress and will continue to evolve based on feedback and additional functionality.
