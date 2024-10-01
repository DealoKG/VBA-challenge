# VBA-challenge
This is a repository for the Miami Data Analytics Bootcamp Module 2 Challenge

# Stock Analysis VBA Script

## Overview

This VBA script performs a stock analysis on data contained within an Excel workbook. It calculates the quarterly change, percentage change, and total stock volume for each ticker symbol. Additionally, it identifies the ticker with the greatest percentage increase, greatest percentage decrease, and greatest total volume.

## Features

- **Quarterly Change Calculation**: Computes the change in stock price from the opening price at the beginning of a quarter to the closing price at the end of the quarter.
- **Percentage Change Calculation**: Computes the percentage change in stock price over the quarter.
- **Total Volume Calculation**: Sums the total stock volume for each ticker over the quarter.
- **Highlighting**: Highlights cells based on the value of the percentage change.
- **Summary Statistics**: Identifies and displays the ticker with the greatest percentage increase, greatest percentage decrease, and greatest total volume.

## How to Use

1. **Prepare Your Data**: Ensure your Excel worksheet contains the following columns:
   - Column A: Ticker
   - Column B: Date
   - Column C: Opening Price
   - Column F: Closing Price
   - Column G: Volume

2. **Run the Script**: Open the VBA editor (Alt + F11), insert a new module, and paste the script into the module. Run the `Analysis` macro.

3. **View Results**: The results will be displayed in the worksheet with the following columns:
   - Column I: Ticker
   - Column J: Quarterly Change
   - Column K: Percentage Change
   - Column L: Total Stock Volume

   Summary statistics will be displayed in columns O to Q.
