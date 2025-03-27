<img title="logo_RSE" src="assets\readme_images\logo_RSE.PNG" alt="logo_RSE" data-align="center" width="400">

---

# CACER Simulator

This repository contains a simulation tool for assessing the **economic**, **financial**, and **energy** performance of renewable energy sharing configurations such as CACER (Configurations for Renewable Energy Sharing in Collective Self-Consumption).

## Description

The simulator supports the evaluation of different collective self-consumption scenarios, including Renewable Energy Communities (RECs) and Groups of Remote Self-Consumers. It provides detailed metrics such as:

- **Economic benefits**: savings and revenues from self-consumed and exported energy.
- **Financial indicators**: Payback Period, Net Present Value (NPV), and Internal Rate of Return (IRR).
- **Energy performance**: self-consumption levels, self-sufficiency, and CO₂ emissions reduction.

## Flow chart CACER simulator

<img title="Flow_chart" src="assets\readme_images\Flow_chart_simulator.png" alt="Flow_chart" data-align="center" width="600">

## Repository Structure

- `assets/`: contains visual or auxiliary resources.
- `files/`: input files and configuration data for simulations.
- `Funzioni_Demand_Side_Management.py`: functions for simulating demand-side flexibility and management.
- `Funzioni_Energy_Model.py`: core energy modeling functions for CACER simulations.
- `Funzioni_Financial_Model.py`: functions for financial analysis and investment evaluation.
- `Funzioni_Generali.py`: general-purpose utility functions used throughout the project.
- `config.yml`: configuration file with key parameters for the simulations.
- `main - CACER tutorial.ipynb`: interactive Jupyter Notebook with step-by-step instructions for using the simulator.
- `reporting_v3.ipynb`: notebook to generate performance reports.
- `reporting_v3.py`: standalone script for report generation.
- `users CACER.xlsx`: example Excel file with user consumption and participation data.

## Prerequisites

You’ll need:

- Python 3.x
- Required libraries listed in `requirements.txt`

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/ToniRollo/CACER-simulator.git
