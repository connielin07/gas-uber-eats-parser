# Uber Eats Receipt Parser

This project uses Google Apps Script to automate receipt data extraction from Gmail and synchronize structured expense records to Google Sheets.  
It also includes a simple dashboard-related structure for viewing and managing parsed expense data.

本專案使用 Google Apps Script 讀取 Gmail 中的 Uber Eats 收據信件，解析日期、金額與相關資訊，並將整理後的資料同步到 Google Sheets，作為 Web 自動化與資料整理練習。

## Project Purpose

The purpose of this project is to practice web-based automation, data parsing, and spreadsheet-based data management.

This project connects to the course **Web Fundamentals and Technologies** by using Google Apps Script, Gmail, Google Sheets, and HTML dashboard-related files.

## Features

- Search Uber Eats receipt emails from Gmail
- Parse receipt content
- Extract structured expense information
- Write data into Google Sheets
- Avoid duplicate records using message-based identification
- Record parsing or processing status
- Dashboard-related HTML and Apps Script controller files
- Support scheduled automation concept

## Tech Stack

- Google Apps Script
- JavaScript
- HTML
- GmailApp
- SpreadsheetApp
- Google Sheets
- clasp

## Data Flow

Gmail Receipt Emails
    ↓
Google Apps Script
    ↓
HTML / Text Parsing
    ↓
Structured Expense Data
    ↓
Google Sheets
    ↓
Dashboard / Statistics View

## Demo

- [Video](https://drive.google.com/file/d/1QcPkBdlXryhT8pE2L7oAT6jvbNndqfB4/view?usp=sharing)

## Docs

- [Report](https://docs.google.com/document/d/1hv5QwfhuxXrQkq5bKuJhLAvPtmU5az3nIVTiSJEvkXs/edit?usp=sharing)
- [Presentation](https://canva.link/q4j2buap12l9wcz)
  
## Course Connection

| Course | Connection |
|---|---|
| Web Fundamentals and Technologies	| Web automation, Apps Script, HTML dashboard |
| Excel / Spreadsheet Statistical Applications | Use spreadsheet as structured data storage and analysis base |
| Data Processing	| Convert semi-structured email content into structured records |

## What I Learned

Through this project, I practiced how to retrieve data from Gmail, parse receipt content, and store the result in a structured spreadsheet format.
I also learned that real-world data is often inconsistent, so error handling, duplicate detection, and data cleaning are important parts of automation.
