This project demonstrates how Python can be used to automate data analysis and generate business-ready Excel reports from raw e-commerce datasets.

# Amazon Product Insights

A Python automation project that transforms raw Amazon product data into a structured Excel analytics report.

The script cleans and processes the dataset, extracts business insights, and automatically generates a professional Excel report containing organized product data, analytical summaries, and visual charts.

This project demonstrates practical skills in data automation, data analysis, and report generation using Python.

---

## Project Overview

Raw datasets are often difficult to interpret directly.
This project automates the process of transforming raw e-commerce product data into a readable and structured business report.

The script performs:

* Data cleaning and preprocessing
* Dataset simplification
* Business insight extraction
* Automated Excel report generation
* Chart creation for visual analysis

The final output is an Excel file containing structured data, insights, and charts that help understand product performance and category distribution.

---

## Technologies Used

* Python
* CSV data processing
* Excel automation with openpyxl
* Basic data analysis techniques

---

## Project Structure

```
project-folder
│
├── app.py
├── amazon.csv
├── final_report.xlsx
└── README.md
```

**app.py**
Main script responsible for data processing and report generation.

**amazon.csv**
Raw dataset containing product information.

**final_report.xlsx**
Automatically generated Excel report with insights and charts.

---

## Features

* Automated data cleaning
* Structured dataset generation
* Business-oriented insights
* Excel report generation
* Automatic chart creation
* Organized and readable output

---

## Example Insights Generated

The report answers questions such as:

* What is the average product rating?
* Which product has the highest number of reviews?
* Which category contains the most products?
* Which category has the highest average rating?
* Which product has the largest discount?
* Which product generates the highest estimated revenue?

---

## How to Run the Project

### 1 Install Python

Make sure Python is installed on your system.

### 2 Install dependencies

```
pip install openpyxl
```

### 3 Place the dataset

Put the dataset file in the project folder:

```
amazon.csv
```

### 4 Run the script

```
python app.py
```

### 5 Generated Output

After running the script, the following file will be created:

```
final_report.xlsx
```

This Excel file contains:

* Product dataset
* Business insights
* Visual charts

---

## Learning Goals

This project was created to practice:

* Data automation
* Data cleaning
* Data analysis
* Python scripting
* Automated Excel reporting

---

## Future Improvements

Possible future improvements include:

* Interactive dashboards
* Additional business metrics
* Advanced visualizations
* Integration with larger datasets

---

## License

This project is open for learning and portfolio purposes.
