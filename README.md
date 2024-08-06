# PowerPoint Presentation Generator

This project generates a PowerPoint presentation from a JSON file using a provided template. The generated presentation can be downloaded directly from the Streamlit app interface.

## Features

- Upload a JSON file containing slide content
- Upload a PowerPoint template file
- Generate a formatted PowerPoint presentation based on the provided template and JSON data
- Download the generated presentation

## Requirements

- Python 3.6 or higher
- Streamlit
- python-pptx

## Installation

1. Clone the repository:

   ```bash
   git clone origin https://github.com/Sour-abh-Raj/PPTGen.git
   cd ppt-generator
   ```

2. Install the required packages:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the Streamlit app:

   ```bash
   streamlit run app.py
   ```

2. Open the provided URL in your browser.

3. Upload the JSON file with the slide content and the PowerPoint template file.

4. Click the "Download PPT" button to download the generated presentation.

## JSON Data Structure

The JSON file should be structured as follows:

```json
[
  {
    "title": "Slide Title 1",
    "content": ["Content line 1", "Content line 2", "Content line 3"]
  },
  {
    "title": "Slide Title 2",
    "content": ["Content line 1", "Content line 2", "Content line 3"]
  }
  // Add more slides as needed
]
```
